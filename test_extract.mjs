function extractLunchName(description) {
  if (!description) return null;
  const upper = description.toUpperCase();
  const originalIndices = [];
  for (let i = 0; i < upper.length; i++) {
    if (upper[i] !== ' ') originalIndices.push(i);
  }
  const spaceless = upper.replace(/\s+/g, '');
  const markerStr = 'CHUYENKHOANLUNCH';
  const markerIdx = spaceless.indexOf(markerStr);
  if (markerIdx === -1) return null;
  const afterIdx = markerIdx + markerStr.length;
  if (afterIdx >= spaceless.length) return null;
  let after = upper.substring(originalIndices[afterIdx]).trim();
  const delimiters = [
    /\.\w/,
    /-\s/,
    /-CHUYEN/i,
    /\s+FT\d/i,
    /\s+CT\s/i,
    /\s+(?=[A-Z]*\d)(?=\d*[A-Z])[A-Z\d]{4,}/i,
    /\s+\d/,
    /\s{2,}Ma\s/i,
    /\.\s+TU:/i,
    /\s+Ma\s+giao/i,
    /\s+Ma\s+GD/i,
    /-\s*$/,
  ];
  let endPos = after.length;
  for (const delim of delimiters) {
    const match = after.match(delim);
    if (match && match.index < endPos) endPos = match.index;
  }
  const nameStr = after.substring(0, endPos).trim();
  const cleanedName = nameStr.replace(/\s+/g, '');
  return cleanedName || null;
}

const cases = [
  { input: 'TRUONG VINH TAN DAT CHUYEN KHOAN LUNCH KEVIN FT26155877 812032 kCMYK4AM/272108', expected: 'KEVIN' },
  { input: 'DO DUY HUNG MBVCB.14514672876.520369.CHUYEN KHO AN LUNCH HARRY.CT tu 0441000649180 DO DUY HUNG toi 0934860931 TA QUOC KHANH tai MB- Ma GD ACSP/ cl520369', expected: 'HARRY' },
  { input: 'Momo chi ho 247 OL 1987501343 260604 CH 13194666304 2 0902969860 CHUYEN KHOAN LUNCH CLI NT- Ma GD ACSP/ co291473', expected: 'CLINT' },
  { input: 'CHUYEN KHOAN LUNCH KANE   Ma giao dich 12345', expected: 'KANE' },
  { input: 'CHUYEN KHOAN LUNCH MIC HAEL-CHUYEN TIEN ABC', expected: 'MICHAEL' },
  { input: 'CHUYEN KHOAN LUNCH ADA M. TU: ZION', expected: 'ADAM' },
  { input: 'HUYNH NGUYEN QUANG TIN CHUYEN KHOAN LUNCH MICHAEL I2A7RXGG /122747', expected: 'MICHAEL' },
];

for (const c of cases) {
  const result = extractLunchName(c.input);
  const pass = result === c.expected;
  console.log(pass ? '✅' : '❌', 'Expected:', c.expected, '| Got:', result);
}
