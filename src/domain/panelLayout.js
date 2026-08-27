export const getSizeType = (rowSpan, colSpan) => {
  if (rowSpan === 1 && colSpan === 1) return '1/16（1コマ）';
  if (rowSpan === 2 && colSpan === 1) return '1/8 縦（2コマ）';
  if (rowSpan === 1 && colSpan === 2) return '1/8 横（2コマ）';
  if (rowSpan === 3 && colSpan === 1) return '3/16 縦（3コマ）';
  if (rowSpan === 1 && colSpan === 3) return '3/16 横（3コマ）';
  if (rowSpan === 2 && colSpan === 2) return '1/4（4コマ）';
  if (rowSpan === 4 && colSpan === 1) return '1/4 縦（4コマ）';
  if (rowSpan === 1 && colSpan === 4) return '1/4 横（4コマ）';
  if (rowSpan === 3 && colSpan === 2) return '6/16 縦（6コマ）';
  if (rowSpan === 2 && colSpan === 3) return '6/16 横（6コマ）';
  if (rowSpan === 4 && colSpan === 2) return '1/2 縦（8コマ）';
  if (rowSpan === 2 && colSpan === 4) return '1/2 横（8コマ）';
  if (rowSpan === 4 && colSpan === 3) return '12/16 縦（12コマ）';
  if (rowSpan === 4 && colSpan === 4) return '1P（16コマ）';
  return 'custom';
};

export const getCoords = (index) => ({
  row: Math.floor(index / 4),
  col: index % 4
});

export const getSpansFromSizeTypeRobust = (sizeType) => {
  if (!sizeType) return { r: 1, c: 1 };

  const normalized = String(sizeType).normalize('NFKC').toLowerCase();
  const value = normalized.replace(/\s+/g, '');
  const hasVertical = /[\u7e26]|vertical/.test(value);
  const hasHorizontal = /[\u6a2a]|horizontal/.test(value);

  if (value.includes('1/16')) return { r: 1, c: 1 };
  if (value.includes('1/8')) return hasVertical ? { r: 2, c: 1 } : { r: 1, c: 2 };
  if (value.includes('3/16')) return hasVertical ? { r: 3, c: 1 } : { r: 1, c: 3 };
  if (value.includes('1/4')) {
    if (hasVertical) return { r: 4, c: 1 };
    if (hasHorizontal) return { r: 1, c: 4 };
    return { r: 2, c: 2 };
  }
  if (value.includes('6/16')) return hasVertical ? { r: 3, c: 2 } : { r: 2, c: 3 };
  if (value.includes('1/2')) return hasVertical ? { r: 4, c: 2 } : { r: 2, c: 4 };
  if (value.includes('12/16')) return hasVertical ? { r: 4, c: 3 } : { r: 3, c: 4 };
  if (value.includes('1p')) return { r: 4, c: 4 };

  const countMatch = value.match(/(\d+)\s*コマ/);
  const count = countMatch ? parseInt(countMatch[1], 10) : NaN;
  if (count === 1) return { r: 1, c: 1 };
  if (count === 2) return hasVertical ? { r: 2, c: 1 } : { r: 1, c: 2 };
  if (count === 3) return hasVertical ? { r: 3, c: 1 } : { r: 1, c: 3 };
  if (count === 4) {
    if (hasVertical) return { r: 4, c: 1 };
    if (hasHorizontal) return { r: 1, c: 4 };
    return { r: 2, c: 2 };
  }
  if (count === 6) return hasVertical ? { r: 3, c: 2 } : { r: 2, c: 3 };
  if (count === 8) return hasVertical ? { r: 4, c: 2 } : { r: 2, c: 4 };
  if (count === 12) return hasVertical ? { r: 4, c: 3 } : { r: 3, c: 4 };
  if (count === 16) return { r: 4, c: 4 };

  return { r: 1, c: 1 };
};

export const canPlacePanelAt = (startIndex, rowSpan, colSpan, occupied) => {
  const startRow = Math.floor(startIndex / 4);
  const startCol = startIndex % 4;
  if (startRow + rowSpan > 4 || startCol + colSpan > 4) return false;

  for (let row = 0; row < rowSpan; row++) {
    for (let col = 0; col < colSpan; col++) {
      const index = (startRow + row) * 4 + (startCol + col);
      if (occupied.has(index)) return false;
    }
  }
  return true;
};

export const findFirstPlaceableIndex = (rowSpan, colSpan, occupied, startIndex = 0) => {
  for (let index = startIndex; index < 16; index++) {
    if (canPlacePanelAt(index, rowSpan, colSpan, occupied)) return index;
  }
  if (startIndex > 0) {
    for (let index = 0; index < startIndex; index++) {
      if (canPlacePanelAt(index, rowSpan, colSpan, occupied)) return index;
    }
  }
  return -1;
};

export const fillPanelArea = (panels, startIndex, rowSpan, colSpan, occupied) => {
  const startRow = Math.floor(startIndex / 4);
  const startCol = startIndex % 4;

  for (let row = 0; row < rowSpan; row++) {
    for (let col = 0; col < colSpan; col++) {
      const index = (startRow + row) * 4 + (startCol + col);
      occupied.add(index);
      if (index !== startIndex) {
        panels[index] = { ...panels[index], hidden: true };
      }
    }
  }
};
