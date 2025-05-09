/* eslint-disable @typescript-eslint/no-explicit-any */
import AdmZip from 'adm-zip';
import * as fs from 'fs';
import { parseStringPromise } from 'xml2js';
import { VisioMaster, VisioPage, VisioRelationship, VisioStylesheet, VisioShape, VisioFile, Style } from './types';

const jsonObjects: any = {};

let masters: VisioMaster[] = [];
let pages: VisioPage[] = [];
let relationships: VisioRelationship[] = [];
let stylesheets: VisioStylesheet[] = [];

export async function parseVisioFile(filePath: string): Promise<VisioFile> {
  const vsdxBuffer = fs.readFileSync(filePath);
  const archive = new AdmZip(vsdxBuffer);

  const entries = archive.getEntries();
  for (const entry of entries) {
    let jsonObj = undefined;

    if (entry.entryName.endsWith('.rels')) {
      const xmlContent = entry.getData().toString('utf-8');
      jsonObj = await parseStringPromise(xmlContent);
      relationships.push.apply(relationships, parseRelationships(jsonObj));
    }

    if (entry.entryName.endsWith('.xml')) {
      const xmlContent = entry.getData().toString('utf-8');
      jsonObj = await parseStringPromise(xmlContent);
      const fileName = getEntryName(entry.entryName);

      switch (fileName) {
        case 'document':
          stylesheets = parseDocumentProperties(jsonObj);
          break;
        case 'masters':
          masters = parseMastersFile(jsonObj);
          break;
        case 'pages':
          pages = parsePagesFile(jsonObj);
          break;
        default:
          jsonObjects[fileName] = jsonObj;
      }
    }
  }

  for (const page of pages) {
    const rel: VisioRelationship | undefined = relationships.find(
      (relation) => relation.Id === page.RelationshipId && relation.Type === 'Page'
    );
    if (rel) {
      const entryName = getEntryName(rel.Target);
      const pageObject = jsonObjects[entryName];
      page.Shapes = getShapes(pageObject);
    }
  }

  return {
    Masters: masters,
    Pages: pages,
    Stylesheets: stylesheets,
    Relationships: relationships
  } as VisioFile;
}

const parseRelationships = (jsonObj: any): VisioRelationship[] => {
  const entries: VisioRelationship[] = [];
  const relObjects = jsonObj['Relationships']['Relationship'];

  for (let i = 0; i < relObjects.length; i++) {
    const relationship = {} as VisioRelationship;

    relationship.Id = relObjects[i]['$']['Id'];
    relationship.Target = relObjects[i]['$']['Target'];
    const type = getRelationshipType(relObjects[i]['$']['Type']);
    if (type) {
      relationship.Type = type;
    }

    entries.push(relationship);
  }

  return entries;
};

const parseDocumentProperties = (jsonObj: any) => {
  const styleSheets: VisioStylesheet[] = [];
  const stylesheetObjects = jsonObj['VisioDocument']['StyleSheets'][0]['StyleSheet'];

  for (let i = 0; i < stylesheetObjects.length; i++) {
    const sheet = {} as VisioStylesheet;

    sheet.ID = stylesheetObjects[i]['$']['ID'];
    sheet.Name = stylesheetObjects[i]['$']['Name'];
    sheet.LineStyleRefId = stylesheetObjects[i]['$']['LineStyle'];
    sheet.FillStyleRefId = stylesheetObjects[i]['$']['FillStyle'];
    sheet.TextStyleRefId = stylesheetObjects[i]['$']['TextStyle'];
    sheet.Style = stylesheetObjects[i]['Cell'];

    styleSheets.push(sheet);
  }

  return styleSheets;
};

const parseMastersFile = (jsonObj: any) => {
  const masters: VisioMaster[] = [];
  const masterObjects = jsonObj['Masters']['Master'];

  for (let i = 0; i < masterObjects.length; i++) {
    const master = {} as VisioMaster;
    master.Id = masterObjects[i]['$']['ID'];
    master.Name = masterObjects[i]['$']['Name'];
    master.UniqueID = masterObjects[i]['$']['UniqueID'];
    master.MasterType = masterObjects[i]['$']['MasterType'];
    master.RelationshipId = masterObjects[i]['Rel'][0]['$']['r:id'];
    master.Hidden = masterObjects[i]['$']['Hidden'];
    master.LineStyleRefId = masterObjects[i]['PageSheet']['LineStyle'];
    master.FillStyleRefId = masterObjects[i]['PageSheet']['FillStyle'];
    master.TextStyleRefId = masterObjects[i]['PageSheet']['TextStyle'];

    masters.push(master);
  }

  return masters;
};

const parsePagesFile = (jsonObj: any) => {
  const pages: VisioPage[] = [];
  const objects = jsonObj['Pages']['Page'];

  for (let i = 0; i < objects.length; i++) {
    const page = {} as VisioPage;
    page.Id = objects[i]['$']['ID'];
    page.Name = objects[i]['$']['Name'];
    page.RelationshipId = objects[i]['Rel'][0]['$']['r:id'];

    pages.push(page);
  }

  return pages;
};

const getShapes = (pageObject: any): VisioShape[] => {
  let shapes = [] as VisioShape[];
  const shapeObjects = pageObject['PageContents']['Shapes'][0];
  const connectObjects = pageObject['PageContents']['Connects'];

  try {
    const shapeCount = shapeObjects['Shape'].length;
    for (let i = 0; i < shapeCount; i++) {
      const shape = { Type: 'unknown', IsEdge: false, Label: '' } as VisioShape;
      const shapeContainer = shapeObjects['Shape'][i];
      const cells = shapeContainer['Cell'];

      shape.Id = shapeContainer['$']['ID'];
      shape.MasterId = shapeContainer['$']['Master'];
      const master = masters.find((master) => master.Id === shape.MasterId);

      if (master) {
        shape.Type = master.Name;
        const masterRel: VisioRelationship | undefined = relationships.find(
          (relation) => relation.Id === master.RelationshipId && relation.Type === 'Master'
        );
        if (masterRel) {
          const masterObj = jsonObjects[getEntryName(masterRel?.Target)];
          if (shape.Type === 'Dynamic connector' && connectObjects) {
            const { fromNode, toNode } = getConnectorNodes(connectObjects[0]['Connect'], shape.Id);
            shape.FromNode = fromNode;
            shape.ToNode = toNode;
            shape.IsEdge = true;
          }
        }
      }

      if (shapeContainer['Text'] && shapeContainer['Text'][0]) {
        shape.Label = shapeContainer['Text'][0]['_'].replace(/\r?\n|\r/g, '').trim();
      }

      shape.Style = getStyleFromObject(cells);

      shapes.push(shape);
    }
  } catch (e) {
    console.log(e);
  }

  return shapes;
};

const getConnectorNodes = (connectObjects: any, shapeId: string) => {
  let fromNode = '';
  let toNode = '';

  try {
    const connects = connectObjects.filter(
      // @ts-ignore
      (connect) => connect['$'].FromSheet === shapeId
    );

    // @ts-ignore
    const from = connects.find((c) => c['$'].FromCell === 'BeginX')['$'];
    // @ts-ignore
    const to = connects.find((c) => c['$'].FromCell === 'EndX')['$'];

    fromNode = from.ToSheet;
    toNode = to.ToSheet;
  } catch (e) {
    console.log(e);
  }

  return { fromNode, toNode };
};

const getStyleFromObject = (cells: any[]): Style => {
  const style = {} as Style;
  const lineWeightInPixels = parseFloat(getValueFromCell(cells, 'LineWeight')) * 96;
  style.LineWeight = lineWeightInPixels;
  style.LineColor = getValueFromCell(cells, 'LineColor');
  style.LinePattern = parseFloat(getValueFromCell(cells, 'LinePattern'));
  style.Rounding = parseFloat(getValueFromCell(cells, 'Rounding'));
  style.BeginArrow = parseFloat(getValueFromCell(cells, 'BeginArrow'));
  style.BeginArrowSize = parseFloat(getValueFromCell(cells, 'BeginArrowSize'));
  style.EndArrow = parseFloat(getValueFromCell(cells, 'EndArrow'));
  style.EndArrowSize = parseFloat(getValueFromCell(cells, 'EndArrowSize'));
  style.LineCap = parseFloat(getValueFromCell(cells, 'LineCap'));
  style.FillForeground = getValueFromCell(cells, 'FillForegnd');
  style.FillBackground = getValueFromCell(cells, 'FillBkgnd');
  style.TextColor = getValueFromCell(cells, 'Color');
  style.FillPattern = parseFloat(getValueFromCell(cells, 'FillPattern'));

  return style;
};

const getValueFromCell = (cells: any, field: string): string => {
  let value = '';

  for (let i = 0; i < cells.length; i++) {
    const cell = cells[i];
    if (cell['$']['N'] === field) {
      value = cell['$']['V'];
    }
  }
  return value;
};

const getRelationshipType = (type: string): 'Master' | 'Page' | undefined => {
  const index = type.lastIndexOf('/') + 1;
  switch (type.substring(index)) {
    case 'master':
      return 'Master';
    case 'page':
      return 'Page';
    default:
      return undefined;
  }
};

const getEntryName = (entryName: string) => {
  const nameStartIndex = entryName.lastIndexOf('/') + 1;
  return entryName.substring(nameStartIndex).replace('.xml', '');
};
