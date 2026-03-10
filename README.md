# @wrdhrd/excel-stream

A zero-memory-leak, multi-tab streaming Excel (.xlsx) compiler for Node.js.
Engineered by wrdhrd to handle massive datasets without crashing the V8 heap.

## Installation

\`\`\`bash
npm install @wrdhrd/excel-stream
\`\`\`

## Usage

\`\`\`typescript
import { WrdhrdExcelStream } from '@wrdhrd/excel-stream';
import \* as fs from 'fs';

async function generate() {
const output = fs.createWriteStream('report.xlsx');
const excelStream = new WrdhrdExcelStream(output);

    excelStream.addSheet("Data");

    // Write headers
    await excelStream.writeRow([
        { value: "ID", style: { bold: true, bgColor: "000000", color: "FFFFFF" } },
        { value: "Amount", style: { bold: true } }
    ]);

    // Stream rows with built-in backpressure protection
    for (let i = 0; i < 1000000; i++) {
        await excelStream.writeRow([{ value: i }, { value: Math.random() * 100 }]);
    }

    await excelStream.commit();

}
generate();
\`\`\`
