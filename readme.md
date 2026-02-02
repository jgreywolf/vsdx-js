vsdx-js is a library for parsing Visio files (\*.vsdx) into javascript objects.

## Installation

```bash
npm install vsdx-js
```

## Usage

### JavaScript (ES Modules)

```javascript
import { parseVisioFile } from 'vsdx-js';

const visioData = await parseVisioFile('path/to/your/file.vsdx');
console.log(visioData);
```

### TypeScript

```typescript
import { parseVisioFile, VisioFile, VisioPage, VisioShape } from 'vsdx-js';

const visioData: VisioFile = await parseVisioFile('path/to/your/file.vsdx');

// Access typed data
visioData.Pages.forEach((page: VisioPage) => {
  console.log(`Page: ${page.Name}`);
  page.Shapes.forEach((shape: VisioShape) => {
    console.log(`  Shape: ${shape.Label} (${shape.Type})`);
  });
});
```

#### JavaScript (ES Modules)
```javascript
import { parseVisioFile } from 'vsdx-js';
import * as fs from 'fs';

// Read file into buffer
const buffer = fs.readFileSync('path/to/your/file.vsdx');
const visioData = await parseVisioFile(buffer);
```
#### TypeScript
```typescript
import { parseVisioFile, VisioFile } from 'vsdx-js';
import * as fs from 'fs';

// Read file into buffer
const buffer: Buffer = fs.readFileSync('path/to/your/file.vsdx');
const visioData: VisioFile = await parseVisioFile(buffer);
```
This is useful when working with files loaded via HTTP requests, streams, or other in-memory sources.

## output

```
[
  {
    Id: "0",
    Name: "Page-1",
    RelationshipId: "rId1",
    Shapes: [
      {
        Type: "Rectangle",
        IsEdge: false,
        Label: "",
        Id: "12",
        MasterId: "2",
        Style: {
          LineWeight: 1.0000000000000002,
          LineColor: "",
          LinePattern: NaN,
          Rounding: NaN,
          BeginArrow: NaN,
          BeginArrowSize: NaN,
          EndArrow: NaN,
          EndArrowSize: NaN,
          LineCap: NaN,
          FillForeground: "",
          FillBackground: "",
          TextColor: "",
          FillPattern: NaN,
        },
      },
      {
        Type: "Rectangle",
        IsEdge: false,
        Label: "",
        Id: "13",
        MasterId: "2",
        Style: {
          LineWeight: 1.0000000000000002,
          LineColor: "",
          LinePattern: NaN,
          Rounding: NaN,
          BeginArrow: NaN,
          BeginArrowSize: NaN,
          EndArrow: NaN,
          EndArrowSize: NaN,
          LineCap: NaN,
          FillForeground: "",
          FillBackground: "",
          TextColor: "",
          FillPattern: NaN,
        },
      },
      {
        Type: "Dynamic connector",
        IsEdge: true,
        Label: "",
        Id: "42",
        MasterId: "97",
        FromNode: "12",
        ToNode: "13",
        Style: {
            LineWeight: NaN,
            LineColor: "",
            LinePattern: NaN,
            Rounding: NaN,
            BeginArrow: NaN,
            BeginArrowSize: NaN,
            EndArrow: NaN,
            EndArrowSize: NaN,
            LineCap: NaN,
            FillForeground: "",
            FillBackground: "",
            TextColor: "",
            FillPattern: NaN,
        },
        }
    ],
  },
]
```
