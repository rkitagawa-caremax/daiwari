export const FREE_LABEL_MAX_CHARACTERS = 80;
export const FREE_LABEL_MAX_COLUMNS = 8;
export const FREE_LABEL_MAX_ROWS = 10;

const normalizeLineEndings = (value) => String(value ?? '').replace(/\r\n?/g, '\n');

const getLineLength = (line) => Array.from(line).length;

export const getFreeLabelTextLayout = (value) => {
  const lines = normalizeLineEndings(value).split('\n');
  const lineLengths = lines.map(getLineLength);
  const rows = lineLengths.reduce(
    (total, length) => total + Math.max(1, Math.ceil(length / FREE_LABEL_MAX_COLUMNS)),
    0
  );
  const columns = Math.min(
    FREE_LABEL_MAX_COLUMNS,
    Math.max(1, ...lineLengths.map((length) => Math.min(FREE_LABEL_MAX_COLUMNS, length)))
  );
  const characterCount = lineLengths.reduce((total, length) => total + length, 0);

  return { characterCount, columns, rows };
};

export const constrainFreeLabelText = (value) => {
  const characters = Array.from(normalizeLineEndings(value));
  let result = '';
  let characterCount = 0;

  for (const character of characters) {
    if (character !== '\n' && characterCount >= FREE_LABEL_MAX_CHARACTERS) break;

    const candidate = result + character;
    if (getFreeLabelTextLayout(candidate).rows > FREE_LABEL_MAX_ROWS) break;

    result = candidate;
    if (character !== '\n') characterCount += 1;
  }

  return result;
};

export const formatFreeLabelText = (value) => {
  const constrained = constrainFreeLabelText(value);
  const rows = [];

  constrained.split('\n').forEach((line) => {
    const characters = Array.from(line);
    if (characters.length === 0) {
      rows.push('');
      return;
    }

    for (let index = 0; index < characters.length; index += FREE_LABEL_MAX_COLUMNS) {
      rows.push(characters.slice(index, index + FREE_LABEL_MAX_COLUMNS).join(''));
    }
  });

  return rows.slice(0, FREE_LABEL_MAX_ROWS).join('\n');
};
