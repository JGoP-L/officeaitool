const fs = require('fs');

const files = [
  'PowerPointAi/DocmeePptClient.vb',
  'PowerPointAi/ThemePptTaskPane.vb',
  'PowerPointAi/Ribbon1.vb',
  'PowerPointAi/ThisAddIn.vb',
  'ShareRibbon/Config/ConfigSettings.vb',
];

function stripComment(line) {
  let inString = false;
  for (let index = 0; index < line.length; index += 1) {
    const ch = line[index];
    if (ch === '"') {
      if (inString && line[index + 1] === '"') {
        index += 1;
      } else {
        inString = !inString;
      }
    }
    if (!inString && ch === "'") {
      return line.slice(0, index);
    }
  }
  return line;
}

function isMethodStart(line) {
  return /^(Public|Private|Friend|Protected)\s+(?:(?:Overrides|Shared|Async|Overloads|NotOverridable|Overridable|MustOverride|Shadows|Partial)\s+)*(Sub|Function)\b/i.test(line) &&
    !/\bEnd\s+(Sub|Function)\b/i.test(line);
}

function isClassStart(line) {
  return /^(Public|Private|Friend|Protected|Partial)\s+.*\bClass\b/i.test(line) &&
    !/\bEnd\s+Class\b/i.test(line);
}

function isBlockIf(line) {
  if (!/^If\s/i.test(line)) return false;
  const thenIndex = line.toLowerCase().lastIndexOf(' then');
  if (thenIndex === -1) return false;
  return line.slice(thenIndex + 5).trim() === '';
}

function push(stack, type, file, lineNumber, text) {
  stack.push({ type, file, lineNumber, text });
}

function pop(stack, expected, file, lineNumber, text) {
  const top = stack.pop();
  if (!top) {
    throw new Error(`${file}:${lineNumber}: unexpected ${text}`);
  }
  if (top.type !== expected) {
    throw new Error(`${file}:${lineNumber}: ${text} closes ${expected}, but top block is ${top.type} from ${top.file}:${top.lineNumber}`);
  }
}

function verifyFile(file) {
  const text = fs.readFileSync(file, 'utf8').replace(/^\uFEFF/, '');
  const stack = [];
  let pendingIf = null;
  const lines = text.split(/\r?\n/);

  lines.forEach((rawLine, index) => {
    const lineNumber = index + 1;
    const line = stripComment(rawLine).trim();
    if (!line) return;

    if (pendingIf) {
      const thenIndex = line.toLowerCase().lastIndexOf(' then');
      if (thenIndex !== -1 && line.slice(thenIndex + 5).trim() === '') {
        push(stack, 'If', file, pendingIf.lineNumber, pendingIf.text);
        pendingIf = null;
        return;
      }
      return;
    }

    if (/^End\s+Class\s*$/i.test(line)) return pop(stack, 'Class', file, lineNumber, line);
    if (/^End\s+Sub\s*$/i.test(line)) {
      const expected = stack.length > 0 && stack[stack.length - 1].type === 'LambdaSub' ? 'LambdaSub' : 'Sub';
      return pop(stack, expected, file, lineNumber, line);
    }
    if (/^End\s+Function\b/i.test(line)) {
      const expected = stack.length > 0 && stack[stack.length - 1].type === 'LambdaFunction' ? 'LambdaFunction' : 'Function';
      if (expected === 'Function' && !/^End\s+Function\s*$/i.test(line)) return;
      return pop(stack, expected, file, lineNumber, line);
    }
    if (/^End\s+If\s*$/i.test(line)) return pop(stack, 'If', file, lineNumber, line);
    if (/^End\s+Try\s*$/i.test(line)) return pop(stack, 'Try', file, lineNumber, line);
    if (/^End\s+Using\s*$/i.test(line)) return pop(stack, 'Using', file, lineNumber, line);
    if (/^End\s+Select\s*$/i.test(line)) return pop(stack, 'Select', file, lineNumber, line);
    if (/^End\s+While\s*$/i.test(line)) return pop(stack, 'While', file, lineNumber, line);
    if (/^End\s+With\s*$/i.test(line)) return pop(stack, 'With', file, lineNumber, line);
    if (/^Next\b/i.test(line)) return pop(stack, 'For', file, lineNumber, line);
    if (/^Loop\b/i.test(line)) return pop(stack, 'Do', file, lineNumber, line);

    if (isClassStart(line)) return push(stack, 'Class', file, lineNumber, line);
    if (isMethodStart(line)) {
      const type = /\bSub\b/i.test(line) ? 'Sub' : 'Function';
      return push(stack, type, file, lineNumber, line);
    }
    if (/\bSub\s*\([^)]*\)\s*$/i.test(line)) return push(stack, 'LambdaSub', file, lineNumber, line);
    if (/\bFunction\s*\([^)]*\)(?:\s+As\b.*)?$/i.test(line)) return push(stack, 'LambdaFunction', file, lineNumber, line);
    if (isBlockIf(line)) return push(stack, 'If', file, lineNumber, line);
    if (/^If\s/i.test(line) && !/\bThen\b/i.test(line)) {
      pendingIf = { lineNumber, text: line };
      return;
    }
    if (/^Try\b/i.test(line)) return push(stack, 'Try', file, lineNumber, line);
    if (/^Using\b/i.test(line)) return push(stack, 'Using', file, lineNumber, line);
    if (/^Select\s+Case\b/i.test(line)) return push(stack, 'Select', file, lineNumber, line);
    if (/^For(\s+Each)?\b/i.test(line)) return push(stack, 'For', file, lineNumber, line);
    if (/^Do\b/i.test(line)) return push(stack, 'Do', file, lineNumber, line);
    if (/^While\b/i.test(line)) return push(stack, 'While', file, lineNumber, line);
    if (/^With\b/i.test(line)) return push(stack, 'With', file, lineNumber, line);
  });

  if (stack.length > 0) {
    const top = stack[stack.length - 1];
    throw new Error(`${file}:${top.lineNumber}: unclosed ${top.type} block: ${top.text}`);
  }
  if (pendingIf) {
    throw new Error(`${file}:${pendingIf.lineNumber}: unfinished multi-line If block: ${pendingIf.text}`);
  }
}

for (const file of files) {
  if (!fs.existsSync(file)) throw new Error(`missing VB file: ${file}`);
  verifyFile(file);
}

console.log('vb block balance checks passed');
