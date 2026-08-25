import struct


FREE = 0xFFFFFFFF
END = 0xFFFFFFFE


class CompoundFile:
    def __init__(self, path):
        self.data = path.read_bytes()
        if self.data[:8] != bytes.fromhex("D0CF11E0A1B11AE1"):
            raise ValueError("不是有效的旧版 Office 复合文档")
        self.sector_size = 1 << struct.unpack_from("<H", self.data, 30)[0]
        self.mini_sector_size = 1 << struct.unpack_from("<H", self.data, 32)[0]
        self.mini_cutoff = struct.unpack_from("<I", self.data, 56)[0]
        fat_ids = list(struct.unpack_from("<109I", self.data, 76))
        fat_ids = [item for item in fat_ids if item not in (FREE, END)]
        self.fat = []
        for sector_id in fat_ids[: struct.unpack_from("<I", self.data, 44)[0]]:
            self.fat.extend(struct.unpack(f"<{self.sector_size // 4}I", self._sector(sector_id)))
        directory = self._chain(struct.unpack_from("<I", self.data, 48)[0], self.fat)
        self.entries = []
        for offset in range(0, len(directory), 128):
            entry = directory[offset : offset + 128]
            if len(entry) < 128:
                continue
            name_length = struct.unpack_from("<H", entry, 64)[0]
            name = entry[: max(0, name_length - 2)].decode("utf-16le", errors="ignore")
            self.entries.append({
                "name": name,
                "type": entry[66],
                "start": struct.unpack_from("<I", entry, 116)[0],
                "size": struct.unpack_from("<Q", entry, 120)[0],
            })
        root = next(item for item in self.entries if item["type"] == 5)
        self.mini_stream = self._chain(root["start"], self.fat)[: root["size"]]
        first_mini_fat = struct.unpack_from("<I", self.data, 60)[0]
        mini_fat_data = self._chain(first_mini_fat, self.fat)
        self.mini_fat = list(struct.unpack(f"<{len(mini_fat_data) // 4}I", mini_fat_data))

    def _sector(self, sector_id):
        start = (sector_id + 1) * self.sector_size
        return self.data[start : start + self.sector_size]

    def _chain(self, start, table):
        result = bytearray()
        seen = set()
        current = start
        while current not in (FREE, END) and current < len(table) and current not in seen:
            seen.add(current)
            result.extend(self._sector(current))
            current = table[current]
        return bytes(result)

    def stream(self, name):
        entry = next(item for item in self.entries if item["name"] == name)
        if entry["size"] >= self.mini_cutoff:
            return self._chain(entry["start"], self.fat)[: entry["size"]]
        result = bytearray()
        seen = set()
        current = entry["start"]
        while current not in (FREE, END) and current < len(self.mini_fat) and current not in seen:
            seen.add(current)
            start = current * self.mini_sector_size
            result.extend(self.mini_stream[start : start + self.mini_sector_size])
            current = self.mini_fat[current]
        return bytes(result[: entry["size"]])


def _clean(value):
    return " ".join(value.replace("\x00", " ").split()).strip()


def extract_slides(path):
    stream = CompoundFile(path).stream("PowerPoint Document")
    slides = []

    def records(start, end, current_slide=None):
        offset = start
        while offset + 8 <= end:
            options, record_type, length = struct.unpack_from("<HHI", stream, offset)
            body_start = offset + 8
            body_end = min(body_start + length, end)
            if body_end < body_start or body_start + length > len(stream):
                break
            version = options & 0xF
            slide = current_slide
            if record_type == 1006:
                slide = []
                slides.append(slide)
            if version == 0xF:
                records(body_start, body_end, slide)
            elif slide is not None and record_type in (4000, 4008, 4026):
                raw = stream[body_start:body_end]
                encoding = "utf-16le" if record_type in (4000, 4026) else "cp1252"
                value = _clean(raw.decode(encoding, errors="ignore"))
                if value and value not in slide:
                    slide.append(value)
            offset = body_start + length

    records(0, len(stream))
    return [parts for parts in slides if parts]


def extract_text(path):
    return "\n".join(f"第{index}页：{'；'.join(parts)}" for index, parts in enumerate(extract_slides(path), 1))
