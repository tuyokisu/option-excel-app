(function (factory) {
  "use strict";

  var api = factory();

  if (typeof module === "object" && module.exports) {
    module.exports = api;
  }

  if (typeof window !== "undefined") {
    window.OptionXls = api;
  }
})(function () {
  "use strict";

  var FREESECT = 0xffffffff;
  var ENDOFCHAIN = 0xfffffffe;
  var FATSECT = 0xfffffffd;
  var DIFSECT = 0xfffffffc;
  var NOSTREAM = 0xffffffff;
  var SIGNATURE = [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1];

  var SECTOR = 512;
  var MINI_SECTOR = 64;
  var MINI_CUTOFF = 4096;
  var HEADER_DIFAT = 109;

  function fail(message) {
    throw new Error(message);
  }

  function broken() {
    fail("この .xls ファイルは壊れているか、対応していない形式です。");
  }

  function toBytes(input) {
    return input instanceof Uint8Array ? input : new Uint8Array(input);
  }

  function viewOf(bytes) {
    return new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  }

  function concat(parts) {
    var total = 0;
    parts.forEach(function (part) {
      total += part.length;
    });
    var out = new Uint8Array(total);
    var offset = 0;
    parts.forEach(function (part) {
      out.set(part, offset);
      offset += part.length;
    });
    return out;
  }

  function utf16le(bytes) {
    var text = "";
    for (var i = 0; i + 1 < bytes.length; i += 2) {
      text += String.fromCharCode(bytes[i] | (bytes[i + 1] << 8));
    }
    return text;
  }

  // ---------------------------------------------------------------------------
  // Compound File Binary (the container around the Excel 97-2003 workbook)
  // ---------------------------------------------------------------------------

  function readCfb(input) {
    var bytes = toBytes(input);
    if (bytes.length < SECTOR) {
      broken();
    }
    for (var s = 0; s < SIGNATURE.length; s += 1) {
      if (bytes[s] !== SIGNATURE[s]) {
        broken();
      }
    }
    var view = viewOf(bytes);
    var sectorShift = view.getUint16(0x1e, true);
    var miniShift = view.getUint16(0x20, true);
    if ((sectorShift !== 9 && sectorShift !== 12) || miniShift !== 6) {
      broken();
    }
    var sectorSize = 1 << sectorShift;
    var perSector = sectorSize / 4;
    var numFat = view.getUint32(0x2c, true);
    var firstDir = view.getUint32(0x30, true);
    var cutoff = view.getUint32(0x38, true);
    var firstMiniFat = view.getUint32(0x3c, true);
    var numMiniFat = view.getUint32(0x40, true);
    var firstDifat = view.getUint32(0x44, true);

    function sectorOffset(id) {
      var offset = (id + 1) * sectorSize;
      if (offset + sectorSize > bytes.length) {
        broken();
      }
      return offset;
    }

    var fatSectors = [];
    for (var h = 0; h < HEADER_DIFAT && fatSectors.length < numFat; h += 1) {
      fatSectors.push(view.getUint32(0x4c + h * 4, true));
    }
    var difat = firstDifat;
    var guard = 0;
    while (fatSectors.length < numFat) {
      if (difat === ENDOFCHAIN || difat === FREESECT || guard > bytes.length / sectorSize) {
        broken();
      }
      var difatOffset = sectorOffset(difat);
      for (var d = 0; d < perSector - 1 && fatSectors.length < numFat; d += 1) {
        fatSectors.push(view.getUint32(difatOffset + d * 4, true));
      }
      difat = view.getUint32(difatOffset + (perSector - 1) * 4, true);
      guard += 1;
    }

    var fat = new Uint32Array(fatSectors.length * perSector);
    fatSectors.forEach(function (id, index) {
      var offset = sectorOffset(id);
      for (var j = 0; j < perSector; j += 1) {
        fat[index * perSector + j] = view.getUint32(offset + j * 4, true);
      }
    });

    function chain(start, table) {
      var ids = [];
      var id = start;
      while (id !== ENDOFCHAIN) {
        if (id >= table.length || ids.length > table.length) {
          broken();
        }
        ids.push(id);
        id = table[id];
      }
      return ids;
    }

    function readChain(start) {
      if (start === ENDOFCHAIN || start === FREESECT) {
        return new Uint8Array(0);
      }
      return concat(
        chain(start, fat).map(function (id) {
          var offset = sectorOffset(id);
          return bytes.subarray(offset, offset + sectorSize);
        })
      );
    }

    var dirBytes = readChain(firstDir);
    var entries = [];
    for (var e = 0; e + 128 <= dirBytes.length; e += 128) {
      var raw = dirBytes.slice(e, e + 128);
      var rawView = viewOf(raw);
      var nameLength = rawView.getUint16(64, true);
      var low = rawView.getUint32(120, true);
      var high = sectorShift === 9 ? 0 : rawView.getUint32(124, true);
      entries.push({
        name: utf16le(raw.subarray(0, Math.max(0, Math.min(64, nameLength) - 2))),
        type: raw[66],
        start: rawView.getUint32(116, true),
        size: low + high * 0x100000000,
        raw: raw,
        data: null,
      });
    }
    if (entries.length === 0 || entries[0].type !== 5) {
      broken();
    }

    var root = entries[0];
    var miniStream = readChain(root.start).subarray(0, root.size);
    var miniFat = new Uint32Array(0);
    if (numMiniFat > 0 && firstMiniFat !== ENDOFCHAIN) {
      var miniFatBytes = readChain(firstMiniFat);
      var miniFatView = viewOf(miniFatBytes);
      miniFat = new Uint32Array(miniFatBytes.length / 4);
      for (var m = 0; m < miniFat.length; m += 1) {
        miniFat[m] = miniFatView.getUint32(m * 4, true);
      }
    }

    entries.forEach(function (entry) {
      if (entry.type !== 2) {
        return;
      }
      if (entry.size === 0) {
        entry.data = new Uint8Array(0);
        return;
      }
      var data;
      if (entry.size < cutoff) {
        data = concat(
          chain(entry.start, miniFat).map(function (id) {
            var offset = id * MINI_SECTOR;
            if (offset + MINI_SECTOR > miniStream.length) {
              broken();
            }
            return miniStream.subarray(offset, offset + MINI_SECTOR);
          })
        );
      } else {
        data = readChain(entry.start);
      }
      if (data.length < entry.size) {
        broken();
      }
      entry.data = data.slice(0, entry.size);
    });

    return { entries: entries };
  }

  function emptyDirEntry() {
    var raw = new Uint8Array(128);
    var view = viewOf(raw);
    view.setUint32(68, NOSTREAM, true);
    view.setUint32(72, NOSTREAM, true);
    view.setUint32(76, NOSTREAM, true);
    return raw;
  }

  function writeCfb(cfb) {
    var entries = cfb.entries;
    var miniParts = [];
    var bigParts = [];
    var miniCount = 0;

    entries.forEach(function (entry) {
      if (entry.type !== 2) {
        return;
      }
      var length = entry.data.length;
      if (length === 0) {
        return;
      }
      if (length < MINI_CUTOFF) {
        var count = Math.ceil(length / MINI_SECTOR);
        miniParts.push({ entry: entry, first: miniCount, count: count });
        miniCount += count;
      } else {
        bigParts.push({ entry: entry, count: Math.ceil(length / SECTOR) });
      }
    });

    var miniStreamBytes = miniCount * MINI_SECTOR;
    var miniStreamSectors = Math.ceil(miniStreamBytes / SECTOR);
    var miniFatSectors = Math.ceil((miniCount * 4) / SECTOR);
    var dirSectors = Math.ceil((entries.length * 128) / SECTOR);
    var bigSectors = 0;
    bigParts.forEach(function (part) {
      bigSectors += part.count;
    });
    var otherSectors = dirSectors + miniFatSectors + miniStreamSectors + bigSectors;

    var fatSectors = 1;
    var difatSectors = 0;
    for (;;) {
      difatSectors =
        fatSectors > HEADER_DIFAT ? Math.ceil((fatSectors - HEADER_DIFAT) / (SECTOR / 4 - 1)) : 0;
      var needed = Math.ceil((otherSectors + fatSectors + difatSectors) / (SECTOR / 4));
      if (needed <= fatSectors) {
        break;
      }
      fatSectors = needed;
    }
    var totalSectors = otherSectors + fatSectors + difatSectors;

    var next = 0;
    function take(count) {
      var start = next;
      next += count;
      return start;
    }
    var fatStart = take(fatSectors);
    var difatStart = take(difatSectors);
    var dirStart = take(dirSectors);
    var miniFatStart = take(miniFatSectors);
    var miniStreamStart = take(miniStreamSectors);
    bigParts.forEach(function (part) {
      part.first = take(part.count);
    });

    var fat = new Uint32Array(fatSectors * (SECTOR / 4)).fill(FREESECT);
    function linkRange(start, count) {
      for (var k = 0; k < count; k += 1) {
        fat[start + k] = k === count - 1 ? ENDOFCHAIN : start + k + 1;
      }
    }
    for (var f = 0; f < fatSectors; f += 1) {
      fat[fatStart + f] = FATSECT;
    }
    for (var x = 0; x < difatSectors; x += 1) {
      fat[difatStart + x] = DIFSECT;
    }
    linkRange(dirStart, dirSectors);
    linkRange(miniFatStart, miniFatSectors);
    linkRange(miniStreamStart, miniStreamSectors);
    bigParts.forEach(function (part) {
      linkRange(part.first, part.count);
    });

    var out = new Uint8Array(SECTOR + totalSectors * SECTOR);
    var view = viewOf(out);
    function sectorAt(id) {
      return SECTOR + id * SECTOR;
    }

    SIGNATURE.forEach(function (byte, index) {
      out[index] = byte;
    });
    view.setUint16(0x18, 0x003e, true);
    view.setUint16(0x1a, 0x0003, true);
    view.setUint16(0x1c, 0xfffe, true);
    view.setUint16(0x1e, 9, true);
    view.setUint16(0x20, 6, true);
    view.setUint32(0x2c, fatSectors, true);
    view.setUint32(0x30, dirStart, true);
    view.setUint32(0x38, MINI_CUTOFF, true);
    view.setUint32(0x3c, miniFatSectors ? miniFatStart : ENDOFCHAIN, true);
    view.setUint32(0x40, miniFatSectors, true);
    view.setUint32(0x44, difatSectors ? difatStart : ENDOFCHAIN, true);
    view.setUint32(0x48, difatSectors, true);
    for (var hd = 0; hd < HEADER_DIFAT; hd += 1) {
      view.setUint32(0x4c + hd * 4, hd < fatSectors ? fatStart + hd : FREESECT, true);
    }
    var perDifat = SECTOR / 4 - 1;
    for (var ds = 0; ds < difatSectors; ds += 1) {
      var base = sectorAt(difatStart + ds);
      for (var di = 0; di < perDifat; di += 1) {
        var fatIndex = HEADER_DIFAT + ds * perDifat + di;
        view.setUint32(base + di * 4, fatIndex < fatSectors ? fatStart + fatIndex : FREESECT, true);
      }
      view.setUint32(
        base + perDifat * 4,
        ds === difatSectors - 1 ? ENDOFCHAIN : difatStart + ds + 1,
        true
      );
    }
    for (var fi = 0; fi < fat.length; fi += 1) {
      view.setUint32(sectorAt(fatStart) + fi * 4, fat[fi], true);
    }

    var miniFat = new Uint32Array(miniFatSectors * (SECTOR / 4)).fill(FREESECT);
    var miniBase = sectorAt(miniStreamStart);
    miniParts.forEach(function (part) {
      for (var k = 0; k < part.count; k += 1) {
        miniFat[part.first + k] = k === part.count - 1 ? ENDOFCHAIN : part.first + k + 1;
      }
      out.set(part.entry.data, miniBase + part.first * MINI_SECTOR);
    });
    for (var mi = 0; mi < miniFat.length; mi += 1) {
      view.setUint32(sectorAt(miniFatStart) + mi * 4, miniFat[mi], true);
    }
    bigParts.forEach(function (part) {
      out.set(part.entry.data, sectorAt(part.first));
    });

    var dirBase = sectorAt(dirStart);
    for (var ei = 0; ei < dirSectors * (SECTOR / 128); ei += 1) {
      var entry = entries[ei];
      var raw = entry ? entry.raw.slice() : emptyDirEntry();
      var rawView = viewOf(raw);
      if (entry && entry.type === 5) {
        rawView.setUint32(116, miniStreamSectors ? miniStreamStart : ENDOFCHAIN, true);
        rawView.setUint32(120, miniStreamBytes, true);
        rawView.setUint32(124, 0, true);
      } else if (entry && entry.type === 2) {
        var part = null;
        miniParts.concat(bigParts).forEach(function (candidate) {
          if (candidate.entry === entry) {
            part = candidate;
          }
        });
        rawView.setUint32(116, part ? part.first : ENDOFCHAIN, true);
        rawView.setUint32(120, entry.data.length, true);
        rawView.setUint32(124, 0, true);
      } else if (entry && entry.type === 1) {
        rawView.setUint32(116, 0, true);
        rawView.setUint32(120, 0, true);
        rawView.setUint32(124, 0, true);
      }
      out.set(raw, dirBase + ei * 128);
    }

    return out;
  }

  // ---------------------------------------------------------------------------
  // BIFF8 records
  // ---------------------------------------------------------------------------

  var RT = {
    FORMULA: 0x0006,
    EOF: 0x000a,
    FILEPASS: 0x002f,
    CONTINUE: 0x003c,
    DEFCOLWIDTH: 0x0055,
    BOUNDSHEET: 0x0085,
    MULRK: 0x00bd,
    MULBLANK: 0x00be,
    RSTRING: 0x00d6,
    DBCELL: 0x00d7,
    SST: 0x00fc,
    LABELSST: 0x00fd,
    EXTSST: 0x00ff,
    RECALCID: 0x01c1,
    DIMENSIONS: 0x0200,
    BLANK: 0x0201,
    NUMBER: 0x0203,
    LABEL: 0x0204,
    BOOLERR: 0x0205,
    STRING: 0x0207,
    ROW: 0x0208,
    INDEX: 0x020b,
    ARRAY: 0x0221,
    DEFAULTROWHEIGHT: 0x0225,
    TABLE: 0x0236,
    RK: 0x027e,
    SHRFMLA: 0x04bc,
    BOF: 0x0809,
  };

  var CELL_TYPES = [
    RT.FORMULA,
    RT.MULRK,
    RT.MULBLANK,
    RT.RSTRING,
    RT.LABELSST,
    RT.BLANK,
    RT.NUMBER,
    RT.LABEL,
    RT.BOOLERR,
    RT.RK,
  ];
  var ATTACHED_TYPES = [RT.STRING, RT.SHRFMLA, RT.ARRAY, RT.TABLE, RT.CONTINUE];
  var MAX_RECORD_DATA = 8224;

  function parseRecords(stream, start, end) {
    var view = viewOf(stream);
    var records = [];
    var pos = start;
    while (pos < end) {
      if (pos + 4 > end) {
        broken();
      }
      var type = view.getUint16(pos, true);
      var length = view.getUint16(pos + 2, true);
      if (pos + 4 + length > end) {
        broken();
      }
      records.push({ type: type, pos: pos, data: stream.subarray(pos + 4, pos + 4 + length) });
      pos += 4 + length;
    }
    return records;
  }

  function recordBytes(type, data) {
    if (data.length > MAX_RECORD_DATA) {
      fail("Excelのデータが大きすぎて書き込めません。");
    }
    var out = new Uint8Array(4 + data.length);
    var view = viewOf(out);
    view.setUint16(0, type, true);
    view.setUint16(2, data.length, true);
    out.set(data, 4);
    return out;
  }

  function u16(data, offset) {
    return data[offset] | (data[offset + 1] << 8);
  }

  function u32(data, offset) {
    return viewOf(data).getUint32(offset, true);
  }

  function shortSheetName(data, offset) {
    var count = data[offset];
    var high = data[offset + 1] & 1;
    var start = offset + 2;
    if (high) {
      return utf16le(data.subarray(start, start + count * 2));
    }
    var text = "";
    for (var i = 0; i < count; i += 1) {
      text += String.fromCharCode(data[start + i]);
    }
    return text;
  }

  function le16(value) {
    return [value & 0xff, (value >> 8) & 0xff];
  }

  function le32(value) {
    return [value & 0xff, (value >>> 8) & 0xff, (value >>> 16) & 0xff, (value >>> 24) & 0xff];
  }

  function setU16(data, offset, value) {
    data[offset] = value & 0xff;
    data[offset + 1] = (value >> 8) & 0xff;
  }

  function setU32(data, offset, value) {
    viewOf(data).setUint32(offset, value >>> 0, true);
  }

  function serialize(records) {
    return concat(
      records.map(function (record) {
        return recordBytes(record.type, record.data);
      })
    );
  }

  function walkDepth(records, visit) {
    var depth = 0;
    records.forEach(function (record, index) {
      if (record.type === RT.BOF) {
        depth += 1;
      }
      visit(record, index, depth);
      if (record.type === RT.EOF) {
        depth -= 1;
      }
    });
  }

  // ---------------------------------------------------------------------------
  // Shared string table
  // ---------------------------------------------------------------------------

  function readSst(group) {
    var pieces = group.map(function (record) {
      return record.data;
    });
    var p = 0;
    var off = 8;
    var total = u32(pieces[0], 0);
    var unique = u32(pieces[0], 4);

    function ensure() {
      while (p < pieces.length && off >= pieces[p].length) {
        p += 1;
        off = 0;
      }
      if (p >= pieces.length) {
        broken();
      }
    }
    function byte() {
      ensure();
      var value = pieces[p][off];
      off += 1;
      return value;
    }
    function skip(count) {
      while (count > 0) {
        ensure();
        var step = Math.min(count, pieces[p].length - off);
        off += step;
        count -= step;
      }
    }

    var strings = [];
    for (var i = 0; i < unique; i += 1) {
      ensure();
      var at = { record: p, offset: off };
      var cch = byte() | (byte() << 8);
      var flags = byte();
      var high = flags & 1;
      var runs = 0;
      var ext = 0;
      if (flags & 8) {
        runs = byte() | (byte() << 8);
      }
      if (flags & 4) {
        ext = (byte() | (byte() << 8) | (byte() << 16) | (byte() << 24)) >>> 0;
      }
      var text = "";
      var remaining = cch;
      while (remaining > 0) {
        if (off >= pieces[p].length) {
          p += 1;
          off = 0;
          if (p >= pieces.length) {
            broken();
          }
          high = pieces[p][off] & 1;
          off += 1;
          continue;
        }
        var width = high ? 2 : 1;
        var fit = Math.floor((pieces[p].length - off) / width);
        if (fit === 0) {
          broken();
        }
        var n = Math.min(fit, remaining);
        for (var k = 0; k < n; k += 1) {
          var at2 = off + k * width;
          text += String.fromCharCode(high ? pieces[p][at2] | (pieces[p][at2 + 1] << 8) : pieces[p][at2]);
        }
        off += n * width;
        remaining -= n;
      }
      skip(runs * 4 + ext);
      strings.push({ text: text, plain: runs === 0 && ext === 0, at: at });
    }
    if (p !== pieces.length - 1 || off !== pieces[p].length) {
      broken();
    }
    return { total: total, unique: unique, strings: strings };
  }

  function appendSst(group, sst, texts) {
    var lastIndex = group.length - 1;
    var current = { index: lastIndex, bytes: Array.prototype.slice.call(group[lastIndex].data) };
    var done = [];
    var positions = [];

    function newRecord() {
      done.push(current);
      current = { index: current.index + 1, bytes: [] };
    }

    texts.forEach(function (text) {
      var high = /[^\u0000-\u00ff]/.test(text) ? 1 : 0;
      var width = high ? 2 : 1;
      if (current.bytes.length + 3 + width > MAX_RECORD_DATA) {
        newRecord();
      }
      positions.push({ record: current.index, offset: current.bytes.length });
      current.bytes.push(text.length & 0xff, (text.length >> 8) & 0xff, high);
      var i = 0;
      while (i < text.length) {
        var room = Math.floor((MAX_RECORD_DATA - current.bytes.length) / width);
        if (room === 0) {
          newRecord();
          current.bytes.push(high);
          continue;
        }
        var n = Math.min(room, text.length - i);
        for (var k = 0; k < n; k += 1) {
          var code = text.charCodeAt(i + k);
          current.bytes.push(code & 0xff);
          if (high) {
            current.bytes.push((code >> 8) & 0xff);
          }
        }
        i += n;
      }
    });
    done.push(current);

    var datas = group.map(function (record) {
      return record.data;
    });
    done.forEach(function (item) {
      datas[item.index] = Uint8Array.from(item.bytes);
    });
    var first = datas[0].slice();
    setU32(first, 4, sst.unique + texts.length);
    datas[0] = first;
    return {
      datas: datas,
      positions: sst.strings
        .map(function (item) {
          return item.at;
        })
        .concat(positions),
    };
  }

  function extSstData(dsst, unique) {
    var buckets = Math.ceil(unique / dsst);
    return new Uint8Array(2 + buckets * 8);
  }

  function pickDsst(original, unique) {
    var dsst = Math.max(8, original || 8);
    while (Math.ceil(unique / dsst) > 128) {
      dsst += 1;
    }
    return dsst;
  }

  // ---------------------------------------------------------------------------
  // Worksheet cell table
  // ---------------------------------------------------------------------------

  function isCellType(type) {
    return CELL_TYPES.indexOf(type) !== -1;
  }

  function isTableType(type) {
    return type === RT.ROW || type === RT.DBCELL || isCellType(type) || ATTACHED_TYPES.indexOf(type) !== -1;
  }

  function colName(col) {
    var n = col + 1;
    var text = "";
    while (n > 0) {
      var rem = (n - 1) % 26;
      text = String.fromCharCode(65 + rem) + text;
      n = Math.floor((n - 1) / 26);
    }
    return text;
  }

  function cellOf(record) {
    var data = record.data;
    var first = u16(data, 2);
    var last = first;
    if (record.type === RT.MULRK || record.type === RT.MULBLANK) {
      last = u16(data, data.length - 2);
    }
    return { type: record.type, data: data, rw: u16(data, 0), colFirst: first, colLast: last, attached: [] };
  }

  function expandMulti(cell) {
    var out = [];
    for (var col = cell.colFirst; col <= cell.colLast; col += 1) {
      var index = col - cell.colFirst;
      var data;
      if (cell.type === RT.MULRK) {
        var o = 4 + index * 6;
        data = Uint8Array.from(le16(cell.rw).concat(le16(col), Array.prototype.slice.call(cell.data.subarray(o, o + 6))));
        out.push({ type: RT.RK, data: data, rw: cell.rw, colFirst: col, colLast: col, attached: [] });
      } else {
        var b = 4 + index * 2;
        data = Uint8Array.from(le16(cell.rw).concat(le16(col), le16(u16(cell.data, b))));
        out.push({ type: RT.BLANK, data: data, rw: cell.rw, colFirst: col, colLast: col, attached: [] });
      }
    }
    return out;
  }

  function hasAttached(cell, types) {
    return cell.attached.some(function (record) {
      return types.indexOf(record.type) !== -1;
    });
  }

  function formulaExpTarget(cell) {
    if (cell.type !== RT.FORMULA || cell.data.length < 27) {
      return null;
    }
    var cce = u16(cell.data, 20);
    if (cce !== 5 || cell.data[22] !== 0x01) {
      return null;
    }
    return { rw: u16(cell.data, 23), col: u16(cell.data, 25) };
  }

  function guardOverwrite(cell, rowsByIndex) {
    var ref = colName(cell.colFirst) + (cell.rw + 1);
    var message =
      ref + " は、ほかのセルと共有している数式（または配列数式）の一部です。上書きすると数式が壊れるため、書き込む行か列を変えてください。";
    if (cell.type !== RT.FORMULA) {
      return;
    }
    if (hasAttached(cell, [RT.SHRFMLA, RT.ARRAY, RT.TABLE])) {
      fail(message);
    }
    var target = formulaExpTarget(cell);
    if (!target) {
      return;
    }
    var row = rowsByIndex[target.rw];
    var master = row
      ? row.cells.filter(function (candidate) {
          return candidate.colFirst === target.col && candidate.type === RT.FORMULA;
        })[0]
      : null;
    if (master && hasAttached(master, [RT.ARRAY, RT.TABLE])) {
      fail(message);
    }
  }

  function readSheet(segment) {
    var records = parseRecords(segment, 0, segment.length);
    if (records.length === 0 || records[0].type !== RT.BOF) {
      broken();
    }
    var info = { records: records, dim: -1, index: -1, defColWidth: -1, defRowHeight: 0x00ff, colInfos: [] };
    walkDepth(records, function (record, index, depth) {
      if (depth !== 1) {
        return;
      }
      if (record.type === RT.DIMENSIONS && info.dim === -1) {
        info.dim = index;
      } else if (record.type === RT.INDEX && info.index === -1) {
        info.index = index;
      } else if (record.type === RT.DEFCOLWIDTH && info.defColWidth === -1) {
        info.defColWidth = index;
      } else if (record.type === RT.DEFAULTROWHEIGHT && record.data.length >= 4) {
        info.defRowHeight = u16(record.data, 2) || 0x00ff;
      } else if (record.type === 0x007d && record.data.length >= 10) {
        info.colInfos.push({ first: u16(record.data, 0), last: u16(record.data, 2), ixfe: u16(record.data, 6) });
      }
    });
    if (info.dim === -1) {
      fail("この .xls のシート構造が読み取れませんでした。Excelで一度上書き保存してから書き込んでください。");
    }

    var tableStart = info.dim + 1;
    var tableEnd = tableStart;
    while (tableEnd < records.length && isTableType(records[tableEnd].type)) {
      tableEnd += 1;
    }
    walkDepth(records, function (record, index, depth) {
      if (index >= tableEnd && depth === 1 && (record.type === RT.ROW || isCellType(record.type))) {
        fail("この .xls のシート構造が読み取れませんでした。Excelで一度上書き保存してから書き込んでください。");
      }
    });

    var rowsByIndex = Object.create(null);
    function rowFor(rw) {
      if (!rowsByIndex[rw]) {
        rowsByIndex[rw] = { rw: rw, rowData: null, cells: [] };
      }
      return rowsByIndex[rw];
    }
    var lastCell = null;
    for (var i = tableStart; i < tableEnd; i += 1) {
      var record = records[i];
      if (record.type === RT.ROW) {
        rowFor(u16(record.data, 0)).rowData = record.data;
        lastCell = null;
      } else if (record.type === RT.DBCELL) {
        lastCell = null;
      } else if (isCellType(record.type)) {
        lastCell = cellOf(record);
        rowFor(lastCell.rw).cells.push(lastCell);
      } else {
        if (!lastCell) {
          broken();
        }
        lastCell.attached.push(record);
      }
    }
    info.tableStart = tableStart;
    info.tableEnd = tableEnd;
    info.rowsByIndex = rowsByIndex;
    return info;
  }

  function defaultIxfe(info, row, col) {
    if (row.rowData && row.rowData.length >= 16 && u16(row.rowData, 12) & 0x0080) {
      return u16(row.rowData, 14) & 0x0fff;
    }
    for (var i = 0; i < info.colInfos.length; i += 1) {
      var colInfo = info.colInfos[i];
      if (col >= colInfo.first && col <= colInfo.last) {
        return colInfo.ixfe;
      }
    }
    return 0x000f;
  }

  function patchRow(info, rw, patches) {
    var rowsByIndex = info.rowsByIndex;
    var row = rowsByIndex[rw];
    if (!row) {
      row = { rw: rw, rowData: null, cells: [] };
      rowsByIndex[rw] = row;
    }
    var removedSst = 0;
    patches.forEach(function (patch) {
      var col = patch.col;
      var cells = [];
      row.cells.forEach(function (cell) {
        if ((cell.type === RT.MULRK || cell.type === RT.MULBLANK) && col >= cell.colFirst && col <= cell.colLast) {
          cells = cells.concat(expandMulti(cell));
        } else {
          cells.push(cell);
        }
      });
      var ixfe = null;
      cells = cells.filter(function (cell) {
        if (cell.colFirst !== col) {
          return true;
        }
        guardOverwrite(cell, rowsByIndex);
        ixfe = u16(cell.data, 4);
        if (cell.type === RT.LABELSST) {
          removedSst += 1;
        }
        return false;
      });
      if (ixfe === null) {
        ixfe = defaultIxfe(info, row, col);
      }
      var head = le16(rw).concat(le16(col), le16(ixfe));
      var next;
      if (patch.kind === "n") {
        var number = new Uint8Array(8);
        viewOf(number).setFloat64(0, Number(patch.value) || 0, true);
        next = { type: RT.NUMBER, data: Uint8Array.from(head.concat(Array.prototype.slice.call(number))) };
      } else {
        next = { type: RT.LABELSST, data: Uint8Array.from(head.concat(le32(patch.isst))) };
      }
      next.rw = rw;
      next.colFirst = col;
      next.colLast = col;
      next.attached = [];
      var at = 0;
      while (at < cells.length && cells[at].colFirst < col) {
        at += 1;
      }
      cells.splice(at, 0, next);
      row.cells = cells;
    });

    var colMic = row.cells[0].colFirst;
    var colMac = row.cells[row.cells.length - 1].colLast + 1;
    var rowData = row.rowData
      ? row.rowData.slice()
      : Uint8Array.from(
          le16(rw).concat(le16(0), le16(0), le16(info.defRowHeight), le16(0), le16(0), le16(0x0100), le16(0x000f))
        );
    setU16(rowData, 2, colMic);
    setU16(rowData, 4, colMac);
    row.rowData = rowData;
    return removedSst;
  }

  function synthesizeRowData(info, row) {
    var colMic = row.cells.length ? row.cells[0].colFirst : 0;
    var colMac = row.cells.length ? row.cells[row.cells.length - 1].colLast + 1 : 0;
    return Uint8Array.from(
      le16(row.rw).concat(le16(colMic), le16(colMac), le16(info.defRowHeight), le16(0), le16(0), le16(0x0100), le16(0x000f))
    );
  }

  function buildCellTable(info) {
    var rows = Object.keys(info.rowsByIndex)
      .map(function (key) {
        return info.rowsByIndex[key];
      })
      .sort(function (a, b) {
        return a.rw - b.rw;
      });
    var parts = [];
    var size = 0;
    var dbcells = [];
    function push(bytes) {
      parts.push(bytes);
      size += bytes.length;
    }
    var i = 0;
    while (i < rows.length) {
      var block = Math.floor(rows[i].rw / 32);
      var blockRows = [];
      while (i < rows.length && Math.floor(rows[i].rw / 32) === block) {
        blockRows.push(rows[i]);
        i += 1;
      }
      var blockStart = size;
      blockRows.forEach(function (row) {
        push(recordBytes(RT.ROW, row.rowData || synthesizeRowData(info, row)));
      });
      var firstCells = [];
      blockRows.forEach(function (row) {
        firstCells.push(row.cells.length ? size : null);
        row.cells.forEach(function (cell) {
          push(recordBytes(cell.type, cell.data));
          cell.attached.forEach(function (record) {
            push(recordBytes(record.type, record.data));
          });
        });
      });
      var dbcellAt = size;
      var dbcell = le32(dbcellAt - blockStart);
      var previous = blockStart + 20;
      firstCells.forEach(function (position) {
        if (position === null) {
          dbcell = dbcell.concat(le16(0));
          return;
        }
        dbcell = dbcell.concat(le16(position - previous));
        previous = position;
      });
      dbcells.push(dbcellAt);
      push(recordBytes(RT.DBCELL, Uint8Array.from(dbcell)));
    }
    return {
      bytes: concat(parts),
      dbcells: dbcells,
      rwMic: rows.length ? rows[0].rw : 0,
      rwMac: rows.length ? rows[rows.length - 1].rw + 1 : 0,
    };
  }

  function widenDimensions(data, rw, colFirst, colLast) {
    var next = data.slice();
    var rwMic = u32(next, 0);
    var rwMac = u32(next, 4);
    var colMic = u16(next, 8);
    var colMac = u16(next, 10);
    if (rwMac === 0 || colMac === 0) {
      rwMic = rw;
      rwMac = rw + 1;
      colMic = colFirst;
      colMac = colLast + 1;
    } else {
      rwMic = Math.min(rwMic, rw);
      rwMac = Math.max(rwMac, rw + 1);
      colMic = Math.min(colMic, colFirst);
      colMac = Math.max(colMac, colLast + 1);
    }
    setU32(next, 0, rwMic);
    setU32(next, 4, rwMac);
    setU16(next, 8, colMic);
    setU16(next, 10, colMac);
    return next;
  }

  function layoutSheet(segment, info, dimensionsData) {
    var records = info.records;
    var table = buildCellTable(info);
    var pre = [];
    for (var i = 0; i < info.tableStart; i += 1) {
      var record = records[i];
      var data = record.data;
      if (i === info.index) {
        data = new Uint8Array(16 + table.dbcells.length * 4);
      } else if (i === info.dim && dimensionsData) {
        data = dimensionsData;
      }
      pre.push({ type: record.type, data: data });
    }
    var defColWidthAt = -1;
    var preSize = 0;
    pre.forEach(function (record, index) {
      if (index === info.defColWidth) {
        defColWidthAt = preSize;
      }
      preSize += 4 + record.data.length;
    });
    var post =
      info.tableEnd < records.length ? segment.subarray(records[info.tableEnd].pos) : new Uint8Array(0);

    return function (sheetStart) {
      if (info.index !== -1) {
        var indexData = pre[info.index].data;
        setU32(indexData, 0, 0);
        setU32(indexData, 4, table.rwMic);
        setU32(indexData, 8, table.rwMac);
        setU32(indexData, 12, defColWidthAt === -1 ? 0 : sheetStart + defColWidthAt);
        table.dbcells.forEach(function (offset, index) {
          setU32(indexData, 16 + index * 4, sheetStart + preSize + offset);
        });
      }
      return concat([serialize(pre), table.bytes, post]);
    };
  }

  function shiftSheet(segment, delta) {
    var copy = segment.slice();
    if (delta === 0) {
      return copy;
    }
    var records = parseRecords(copy, 0, copy.length);
    walkDepth(records, function (record, index, depth) {
      if (depth !== 1 || record.type !== RT.INDEX || record.data.length < 16) {
        return;
      }
      var data = record.data;
      var ibXf = u32(data, 12);
      if (ibXf) {
        setU32(data, 12, ibXf + delta);
      }
      for (var o = 16; o + 4 <= data.length; o += 4) {
        setU32(data, o, u32(data, o) + delta);
      }
    });
    return copy;
  }

  // ---------------------------------------------------------------------------
  // Workbook stream
  // ---------------------------------------------------------------------------

  function pickSheet(sheets) {
    var worksheets = sheets.filter(function (sheet) {
      return sheet.isWorksheet;
    });
    if (worksheets.length === 0) {
      fail("この .xls には書き込めるワークシートがありません。");
    }
    var named = worksheets.filter(function (sheet) {
      return sheet.name === "計算" || sheet.name === "一覧";
    })[0];
    if (named) {
      return named;
    }
    return (
      worksheets.filter(function (sheet) {
        return sheet.name !== "マスタ";
      })[0] || worksheets[0]
    );
  }

  function rewriteWorkbookStream(stream, row, patches, options) {
    var settings = options || {};
    var firstBof = parseRecords(stream, 0, Math.min(stream.length, 20))[0];
    if (!firstBof || firstBof.type !== RT.BOF || u16(firstBof.data, 0) !== 0x0600 || u16(firstBof.data, 2) !== 0x0005) {
      fail("Excel 97-2003 形式（BIFF8）の .xls だけに対応しています。Excelで一度保存し直してください。");
    }

    var globals = [];
    var view = viewOf(stream);
    var pos = 0;
    for (;;) {
      if (pos + 4 > stream.length) {
        broken();
      }
      var type = view.getUint16(pos, true);
      var length = view.getUint16(pos + 2, true);
      globals.push({ type: type, pos: pos, data: stream.subarray(pos + 4, pos + 4 + length) });
      pos += 4 + length;
      if (type === RT.EOF) {
        break;
      }
    }
    var globalsEnd = pos;

    var sheets = [];
    var sstIndex = -1;
    globals.forEach(function (record, index) {
      if (record.type === RT.FILEPASS) {
        fail("パスワードで保護された .xls には書き込めません。保護を外してから書き込んでください。");
      }
      if (record.type === RT.BOUNDSHEET) {
        sheets.push({
          record: index,
          pos: u32(record.data, 0),
          dt: record.data[5],
          name: shortSheetName(record.data, 6),
        });
      }
      if (record.type === RT.SST && sstIndex === -1) {
        sstIndex = index;
      }
    });
    if (sheets.length === 0) {
      broken();
    }

    var ordered = sheets.slice().sort(function (a, b) {
      return a.pos - b.pos;
    });
    if (ordered[0].pos < globalsEnd) {
      broken();
    }
    ordered.forEach(function (sheet, index) {
      sheet.end = index + 1 < ordered.length ? ordered[index + 1].pos : stream.length;
      sheet.segment = stream.subarray(sheet.pos, sheet.end);
      var bof = parseRecords(sheet.segment, 0, Math.min(sheet.segment.length, 20))[0];
      if (!bof || bof.type !== RT.BOF) {
        broken();
      }
      sheet.isWorksheet = sheet.dt === 0 && u16(bof.data, 2) === 0x0010;
    });

    var target = pickSheet(sheets);
    var globalsTail = stream.subarray(globalsEnd, ordered[0].pos);

    var newGlobals = globals.map(function (record) {
      return { type: record.type, data: record.data };
    });
    sheets.forEach(function (sheet) {
      sheet.globalRecord = newGlobals[sheet.record];
    });

    var textPatches = patches.filter(function (patch) {
      return patch.kind !== "n";
    });
    var sstGroupLength = 0;
    var sstPositions = null;
    var sstDsst = 8;
    var sstUnique = 0;
    var sstTotal = 0;
    var extSstIndex = -1;
    if (textPatches.length || sstIndex !== -1) {
      if (sstIndex === -1) {
        fail("この .xls には文字列表がありません。Excelで文字を1つ入力して保存し直してから書き込んでください。");
      }
      var groupEnd = sstIndex + 1;
      while (groupEnd < globals.length && globals[groupEnd].type === RT.CONTINUE) {
        groupEnd += 1;
      }
      var group = globals.slice(sstIndex, groupEnd);
      var sst = readSst(group);
      var lookup = Object.create(null);
      sst.strings.forEach(function (item, index) {
        if (item.plain && !(item.text in lookup)) {
          lookup[item.text] = index;
        }
      });
      var added = [];
      textPatches.forEach(function (patch) {
        var text = String(patch.value);
        if (text.length > 32767) {
          fail("書き込む文字が長すぎます。");
        }
        if (!(text in lookup)) {
          lookup[text] = sst.unique + added.length;
          added.push(text);
        }
        patch.isst = lookup[text];
      });
      var appended = appendSst(group, sst, added);
      var replacement = appended.datas.map(function (data, index) {
        return { type: index === 0 ? RT.SST : RT.CONTINUE, data: data };
      });
      newGlobals.splice.apply(newGlobals, [sstIndex, groupEnd - sstIndex].concat(replacement));
      sstGroupLength = replacement.length;
      sstPositions = appended.positions;
      sstUnique = sst.unique + added.length;
      sstTotal = sst.total + textPatches.length;

      extSstIndex = sstIndex + sstGroupLength;
      if (newGlobals[extSstIndex] && newGlobals[extSstIndex].type === RT.EXTSST) {
        sstDsst = u16(newGlobals[extSstIndex].data, 0);
      } else {
        newGlobals.splice(extSstIndex, 0, { type: RT.EXTSST, data: null });
      }
      sstDsst = pickDsst(sstDsst, sstUnique);
      newGlobals[extSstIndex] = { type: RT.EXTSST, data: extSstData(sstDsst, sstUnique) };
    }

    if (!settings.keepRecalcId) {
      newGlobals.forEach(function (record) {
        if (record.type === RT.RECALCID && record.data.length >= 8) {
          var data = record.data.slice();
          setU32(data, 4, 0);
          record.data = data;
        }
      });
    }

    var info = readSheet(target.segment);
    var rw = row - 1;
    var removedSst = 0;
    var dimensionsData = null;
    if (patches.length) {
      removedSst = patchRow(info, rw, patches);
      var cols = patches.map(function (patch) {
        return patch.col;
      });
      dimensionsData = widenDimensions(
        info.records[info.dim].data,
        rw,
        Math.min.apply(null, cols),
        Math.max.apply(null, cols)
      );
    }
    var buildTarget = layoutSheet(target.segment, info, dimensionsData);

    if (sstPositions) {
      var first = newGlobals[sstIndex].data.slice();
      setU32(first, 0, Math.max(0, sstTotal - removedSst));
      newGlobals[sstIndex].data = first;
    }

    var globalsSize = 0;
    var recordStarts = [];
    newGlobals.forEach(function (record) {
      recordStarts.push(globalsSize);
      globalsSize += 4 + record.data.length;
    });
    globalsSize += globalsTail.length;

    if (sstPositions) {
      var ext = newGlobals[extSstIndex].data;
      setU16(ext, 0, sstDsst);
      for (var bucket = 0; bucket * sstDsst < sstUnique; bucket += 1) {
        var place = sstPositions[bucket * sstDsst];
        var recordStart = recordStarts[sstIndex + place.record];
        var inRecord = 4 + place.offset;
        setU32(ext, 2 + bucket * 8, recordStart + inRecord);
        setU16(ext, 2 + bucket * 8 + 4, inRecord);
        setU16(ext, 2 + bucket * 8 + 6, 0);
      }
    }

    var cursor = globalsSize;
    var sheetBytes = [];
    ordered.forEach(function (sheet) {
      var start = cursor;
      var bytes = sheet === target ? buildTarget(start) : shiftSheet(sheet.segment, start - sheet.pos);
      sheet.newPos = start;
      sheetBytes.push(bytes);
      cursor += bytes.length;
    });
    sheets.forEach(function (sheet) {
      var data = sheet.globalRecord.data.slice();
      setU32(data, 0, sheet.newPos);
      sheet.globalRecord.data = data;
    });

    return concat([serialize(newGlobals), globalsTail].concat(sheetBytes));
  }

  function findWorkbookEntry(cfb) {
    var entry = cfb.entries.filter(function (item) {
      return item.type === 2 && item.name.toLowerCase() === "workbook";
    })[0];
    if (!entry) {
      var old = cfb.entries.some(function (item) {
        return item.type === 2 && item.name.toLowerCase() === "book";
      });
      fail(
        old
          ? "Excel 5.0/95 形式の .xls には対応していません。Excelで『Excel 97-2003 ブック』か .xlsx で保存し直してください。"
          : "この .xls の中にブックが見つかりませんでした。"
      );
    }
    return entry;
  }

  function writeXlsRow(input, row, patches, options) {
    var rowNum = Number(row);
    if (!Number.isInteger(rowNum) || rowNum < 1 || rowNum > 65536) {
      fail("行番号は 1〜65536 の整数にしてください。");
    }
    (patches || []).forEach(function (patch) {
      if (!Number.isInteger(patch.col) || patch.col < 0 || patch.col > 255) {
        fail("書き込む列が正しくありません。");
      }
    });
    var cfb = readCfb(input);
    var entry = findWorkbookEntry(cfb);
    entry.data = rewriteWorkbookStream(entry.data, rowNum, (patches || []).slice(), options);
    return writeCfb(cfb);
  }

  return {
    readCfb: readCfb,
    writeCfb: writeCfb,
    parseRecords: parseRecords,
    shortSheetName: shortSheetName,
    rewriteWorkbookStream: rewriteWorkbookStream,
    findWorkbookEntry: findWorkbookEntry,
    readSst: readSst,
    writeXlsRow: writeXlsRow,
    RT: RT,
    CELL_TYPES: CELL_TYPES,
    ATTACHED_TYPES: ATTACHED_TYPES,
    u16: u16,
    u32: u32,
  };
});
