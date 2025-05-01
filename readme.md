vsdx-js is a library for parsing Visio files (*.vsdx) into javascript objects.

## usage


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