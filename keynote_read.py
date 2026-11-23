"""
Reads slides out of Keynote (.key) and PowerPoint (.pptx) files, without Keynote or PowerPoint.

  slides = read_deck(path)   ->  [Slide(number, texts=[...], images=[MediaRef], media=[MediaRef])]

Keynote 6+ (2013 onward) stores slides as compressed protobuf "IWA" objects inside the package; this module
decodes them generically (snappy + protobuf wire format, no schema files) and follows the show's slide list
to get slide order, the text on each slide, and the pictures/clips placed on it. Older Keynote '09 packages
(index.apxl XML) are read on a best-effort basis. A .key can be a folder (package) or a single zipped file.
"""
import gzip
import io
import os
import posixpath
import re
import zipfile
import xml.etree.ElementTree as ET
from dataclasses import dataclass, field
from pathlib import Path


@dataclass
class MediaRef:
    name: str          # file name inside the deck
    size: int          # bytes
    read: object       # callable -> bytes

    @property
    def ext(self):
        return Path(self.name).suffix.lower().lstrip(".")


@dataclass
class Slide:
    number: int
    texts: list = field(default_factory=list)
    images: list = field(default_factory=list)
    media: list = field(default_factory=list)


class DeckError(Exception):
    pass


AUDIO_VIDEO = {"mp3", "m4a", "wav", "aac", "aif", "aiff", "caf", "mp4", "mov", "m4v", "avi"}


# ------------------------------------------------------------------------------ low-level decoding
def _snappy(b):
    i = n = s = 0
    while True:
        c = b[i]
        i += 1
        n |= (c & 0x7F) << s
        s += 7
        if c < 0x80:
            break
    out = bytearray()
    while i < len(b):
        t = b[i]
        i += 1
        k = t & 3
        if k == 0:
            ln = t >> 2
            if ln >= 60:
                nb = ln - 59
                ln = int.from_bytes(b[i:i + nb], "little")
                i += nb
            ln += 1
            out += b[i:i + ln]
            i += ln
            continue
        if k == 1:
            ln = ((t >> 2) & 7) + 4
            off = ((t >> 5) << 8) | b[i]
            i += 1
        elif k == 2:
            ln = (t >> 2) + 1
            off = int.from_bytes(b[i:i + 2], "little")
            i += 2
        else:
            ln = (t >> 2) + 1
            off = int.from_bytes(b[i:i + 4], "little")
            i += 4
        for _ in range(ln):
            out.append(out[-off])
    return bytes(out)


def _iwa(data):
    i, out = 0, bytearray()
    while i < len(data):
        ln = int.from_bytes(data[i + 1:i + 4], "little")
        out += _snappy(data[i + 4:i + 4 + ln]) if data[i] == 0 else data[i + 4:i + 4 + ln]
        i += 4 + ln
    return bytes(out)


def _varint(b, i):
    n = s = 0
    while True:
        c = b[i]
        i += 1
        n |= (c & 0x7F) << s
        s += 7
        if c < 0x80:
            return n, i


def _fields(b):
    i, out = 0, []
    while i < len(b):
        key, i = _varint(b, i)
        f, w = key >> 3, key & 7
        if f == 0:
            raise ValueError("bad field")
        if w == 0:
            v, i = _varint(b, i)
        elif w == 1:
            v, i = b[i:i + 8], i + 8
        elif w == 5:
            v, i = b[i:i + 4], i + 4
        elif w == 2:
            ln, i = _varint(b, i)
            v, i = b[i:i + ln], i + ln
            if i > len(b):
                raise ValueError("truncated")
        else:
            raise ValueError("bad wire type")
        out.append((f, w, v))
    return out


def _safe_fields(b):
    try:
        return _fields(b)
    except (ValueError, IndexError):
        return []


def _objects(raw):
    i = 0
    while i < len(raw):
        ln, i = _varint(raw, i)
        info = _safe_fields(raw[i:i + ln])
        i += ln
        oid = next((v for f, w, v in info if f == 1 and w == 0), None)
        for f, w, v in info:
            if f == 2 and w == 2:
                mi = _safe_fields(v)
                typ = next((x for ff, ww, x in mi if ff == 1 and ww == 0), None)
                n = next((x for ff, ww, x in mi if ff == 3 and ww == 0), 0)
                if oid is not None:
                    yield oid, typ, raw[i:i + n]
                i += n


# ----------------------------------------------------------------------------------------- packages
class _Package:
    """Uniform access to a .key folder package or a single-file (zipped) .key."""

    def __init__(self, path):
        self.path = Path(path)
        self.zip = None
        if self.path.is_dir():
            self.names = [str(p.relative_to(self.path)).replace(os.sep, "/") for p in self.path.rglob("*") if p.is_file()]
        elif zipfile.is_zipfile(self.path):
            self.zip = zipfile.ZipFile(self.path)
            self.names = self.zip.namelist()
        else:
            raise DeckError("not a Keynote package")
        # Some packages keep Index/*.iwa inside an Index.zip
        self.inner = None
        if "Index.zip" in self.names:
            self.inner = zipfile.ZipFile(io.BytesIO(self.read("Index.zip")))

    def read(self, name):
        if self.zip:
            return self.zip.read(name)
        return (self.path / name).read_bytes()

    def size(self, name):
        if self.zip:
            return self.zip.getinfo(name).file_size
        return (self.path / name).stat().st_size

    def iwa_files(self):
        if self.inner:
            return [(n, self.inner.read(n)) for n in self.inner.namelist() if n.endswith(".iwa")]
        return [(n, self.read(n)) for n in self.names if n.startswith("Index/") and n.endswith(".iwa")]

    def data_files(self):
        return [n for n in self.names if n.startswith("Data/") and not n.endswith("/")]


SKIP_IMAGES = re.compile(r"(^|/)(st-|mt-|PresetImageFill|bullet_)|-small-|-small\.", re.I)


def _read_iwa_deck(pkg):
    objs = {}
    for _, data in pkg.iwa_files():
        try:
            raw = _iwa(data)
        except Exception:
            continue
        for oid, typ, payload in _objects(raw):
            objs[oid] = (typ, payload)
    if not objs:
        raise DeckError("no slide data found")

    def refs(b, depth=0):
        fs = _safe_fields(b)
        if len(fs) == 1 and fs[0][0] == 1 and fs[0][1] == 0 and fs[0][2] in objs:
            return [fs[0][2]]
        out = []
        for f, w, v in fs:
            if w == 2 and depth < 6:
                out += refs(v, depth + 1)
        return out

    def varints(b, depth=0):
        out = []
        for f, w, v in _safe_fields(b):
            if w == 0:
                out.append(v)
            elif w == 2 and depth < 5:
                out += varints(v, depth + 1)
        return out

    # data files are named "<name>-<id>.<ext>"
    data_by_id = {}
    for n in pkg.data_files():
        m = re.search(r"-(\d+)\.[A-Za-z0-9]+$", n)
        if m and not SKIP_IMAGES.search(n):
            data_by_id[int(m.group(1))] = n

    show = next((o for o, (t, _) in objs.items() if t == 2), None)
    if show is None:
        raise DeckError("no slide list found")
    nodes = [r for f, w, v in _safe_fields(objs[show][1]) if f == 3 and w == 2 for r in refs(v)]
    stop = {1, 2, 4, 6, 401}
    slides = []
    for number, node in enumerate(nodes, 1):
        sid = next((r for r in refs(objs[node][1]) if objs[r][0] == 5), None)
        s = Slide(number)
        if sid is None:
            slides.append(s)
            continue
        seen, queue = set(), [(sid, 0)]
        found = []
        while queue:
            oid, d = queue.pop(0)
            if oid in seen or oid not in objs:
                continue
            seen.add(oid)
            t, payload = objs[oid]
            if t == 2001:  # text storage
                try:
                    txt = "".join(v.decode("utf8") for f, w, v in _fields(payload) if f == 3 and w == 2)
                except (ValueError, UnicodeDecodeError, IndexError):
                    txt = ""
                txt = txt.replace("￼", "").replace(" ", "\n").strip()
                if txt:
                    s.texts.append(txt)
                continue
            if t is not None and 3000 <= t < 3100 and t not in (3006, 3016):  # drawables: images, movies, audio
                for v in varints(payload):
                    if v > 100 and v in data_by_id:
                        found.append(data_by_id[v])
            if d >= 4 or (t in stop and oid != sid) or (t == 5 and oid != sid) or (t is not None and 2020 <= t <= 2030):
                continue
            queue += [(r, d + 1) for r in refs(payload)]
        for n in dict.fromkeys(found):
            ref = MediaRef(n.split("/")[-1], pkg.size(n), (lambda nm=n: pkg.read(nm)))
            (s.media if ref.ext in AUDIO_VIDEO else s.images).append(ref)
        slides.append(s)
    return slides


def _read_apxl_deck(pkg):
    """Keynote '09 and earlier: index.apxl (XML, sometimes gzipped). Best effort."""
    name = next((n for n in pkg.names if n.lower() in ("index.apxl", "index.apxl.gz")), None)
    if not name:
        raise DeckError("not a recognised Keynote format")
    data = pkg.read(name)
    if name.endswith(".gz"):
        data = gzip.decompress(data)
    root = ET.fromstring(data)
    local = lambda tag: tag.rsplit("}", 1)[-1]
    slides = []
    for el in root.iter():
        if local(el.tag) != "slide":
            continue
        s = Slide(len(slides) + 1)
        for body in el.iter():
            if local(body.tag) == "text-body":
                txt = "\n".join("".join(p.itertext()).strip() for p in body.iter() if local(p.tag) == "p").strip()
                if txt:
                    s.texts.append(txt)
            if local(body.tag) == "data":
                path = next((v for k, v in body.attrib.items() if local(k) == "path"), "")
                if path and path in pkg.names and not SKIP_IMAGES.search(path):
                    ref = MediaRef(path.split("/")[-1], pkg.size(path), (lambda nm=path: pkg.read(nm)))
                    (s.media if ref.ext in AUDIO_VIDEO else s.images).append(ref)
        slides.append(s)
    if not slides:
        raise DeckError("no slides found in index.apxl")
    return slides


def read_key(path):
    pkg = _Package(path)
    if pkg.inner is not None or any(n.startswith("Index/") and n.endswith(".iwa") for n in pkg.names):
        return _read_iwa_deck(pkg)
    return _read_apxl_deck(pkg)


def read_pptx(path):
    """Reads a .pptx in place: only the slide XML is read. Pictures and clips are listed with their size from the
    zip directory and read only if they're imported (python-pptx would load every video into memory)."""
    from lxml import etree
    ns = {"p": "http://schemas.openxmlformats.org/presentationml/2006/main",
          "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
          "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships"}
    rel_ns = "{http://schemas.openxmlformats.org/package/2006/relationships}"
    r_embed, r_link = f"{{{ns['r']}}}embed", f"{{{ns['r']}}}link"
    z = zipfile.ZipFile(path)
    names = set(z.namelist())
    parser = etree.XMLParser(resolve_entities=False, no_network=True, huge_tree=False)

    def xml(name):
        return etree.fromstring(z.read(name), parser)

    def rels(part):
        d, base = posixpath.split(part)
        rp = posixpath.join(d, "_rels", base + ".rels")
        out = {}
        if rp in names:
            for r in xml(rp).iter(rel_ns + "Relationship"):
                if r.get("TargetMode") == "External":
                    continue
                out[r.get("Id")] = posixpath.normpath(posixpath.join(d, r.get("Target", "")))
        return out

    def ref(part, prefix):
        info = z.getinfo(part)
        ext = posixpath.splitext(part)[1].lower()
        # named by the checksum stored in the zip directory: same picture -> same name, without reading it
        return MediaRef(f"{prefix}-{info.CRC:08x}-{info.file_size}{ext}", info.file_size, (lambda nm=part: z.read(nm)))

    pres = "ppt/presentation.xml"
    prels = rels(pres)
    order = [prels.get(x.get(r_embed) or x.get(f"{{{ns['r']}}}id"))
             for x in xml(pres).iterfind("p:sldIdLst/p:sldId", ns)]
    slides = []
    for number, part in enumerate([o for o in order if o in names], 1):
        s = Slide(number)
        srels = rels(part)
        tree = xml(part).find("p:cSld/p:spTree", ns)
        shapes = [el for el in tree if etree.QName(el).localname in ("sp", "pic", "graphicFrame", "grpSp")]

        def pos(el):
            off = el.find(".//a:off", ns)
            return (int(off.get("y", 0)), int(off.get("x", 0))) if off is not None else (0, 0)
        for el in sorted(shapes, key=pos):
            tag = etree.QName(el).localname
            if tag == "sp":
                body = el.find("p:txBody", ns)
                if body is not None:
                    paras = ["".join("\v" if etree.QName(n).localname == "br" else (n.text or "")
                                     for n in p.iter(f"{{{ns['a']}}}t", f"{{{ns['a']}}}br"))
                             for p in body.iterfind("a:p", ns)]
                    t = "\n".join(paras).strip()
                    if t:
                        s.texts.append(t)
            media_rid = None
            for node in el.iter():
                if isinstance(node.tag, str) and etree.QName(node).localname in ("media", "videoFile", "audioFile"):
                    media_rid = node.get(r_embed) or node.get(r_link)
                    if media_rid:
                        break
            if media_rid and srels.get(media_rid) in names:
                s.media.append(ref(srels[media_rid], "media"))
                continue
            if tag == "pic":
                blip = el.find(".//a:blip", ns)
                target = srels.get(blip.get(r_embed)) if blip is not None else None
                if target in names:
                    s.images.append(ref(target, "image"))
        slides.append(s)
    return slides


def read_deck(path):
    p = Path(path)
    if p.suffix.lower() == ".pptx":
        return read_pptx(p)
    if p.suffix.lower() == ".key":
        return read_key(p)
    raise DeckError(f"not a .key or .pptx file: {p.name}")
