/* eslint-disable @typescript-eslint/no-explicit-any */
import AdmZip from 'adm-zip';
import * as fs from 'fs';
import { parseStringPromise } from 'xml2js';
import { VisioMaster, VisioPage, VisioRelationship, VisioStylesheet, VisioShape, VisioFile, Style } from './types';

// Export all types for consumers
export * from './types';

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

  // Enhanced: Parse master objects and their styles
  for (const master of masters) {
    const rel: VisioRelationship | undefined = relationships.find(
      (relation) => relation.Id === master.RelationshipId && relation.Type === 'Master'
    );
    if (rel) {
      const entryName = getEntryName(rel.Target);
      const masterObject = jsonObjects[entryName];
      if (masterObject) {
        // Parse master's own style properties from its PageSheet
        const pageSheet = masterObject?.PageContents?.PageSheet;
        if (pageSheet && pageSheet.Cell) {
          master.Style = processStyleCells(pageSheet.Cell);
        } else {
          master.Style = {};
        }
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

const processStyleCells = (cells: any[]): Style => {
  const style = {} as Style;
  
  if (!cells || !Array.isArray(cells)) return style;
  
  cells.forEach(cell => {
    const name = cell.$?.N;
    const value = cell.$?.V;
    const unit = cell.$?.U;
    const formula = cell.$?.F;
    
    if (!name || value === undefined) return;
    
    // Skip inherited and themed values for now
    if (formula === 'Inh' || value === 'Themed') return;
    
    switch (name) {
      case 'LineWeight':
        style.LineWeight = parseLineWeight(value, unit);
        break;
      case 'LineColor':
        const lineColor = parseColorEnhanced(value);
        if (lineColor) style.LineColor = lineColor;
        break;
      case 'LinePattern':
        const linePattern = parseFloat(value);
        if (!isNaN(linePattern)) style.LinePattern = linePattern;
        break;
      case 'Rounding':
        const rounding = parseFloat(value);
        if (!isNaN(rounding)) style.Rounding = rounding;
        break;
      case 'BeginArrow':
        const beginArrow = parseInt(value);
        if (!isNaN(beginArrow)) style.BeginArrow = beginArrow;
        break;
      case 'BeginArrowSize':
        const beginArrowSize = parseInt(value);
        if (!isNaN(beginArrowSize)) style.BeginArrowSize = beginArrowSize;
        break;
      case 'EndArrow':
        const endArrow = parseInt(value);
        if (!isNaN(endArrow)) style.EndArrow = endArrow;
        break;
      case 'EndArrowSize':
        const endArrowSize = parseInt(value);
        if (!isNaN(endArrowSize)) style.EndArrowSize = endArrowSize;
        break;
      case 'LineCap':
        const lineCap = parseInt(value);
        if (!isNaN(lineCap)) style.LineCap = lineCap;
        break;
      case 'FillForegnd':
        const fillForeground = parseColorEnhanced(value);
        if (fillForeground) style.FillForeground = fillForeground;
        break;
      case 'FillBkgnd':
        const fillBackground = parseColorEnhanced(value);
        if (fillBackground) style.FillBackground = fillBackground;
        break;
      case 'FillPattern':
        const fillPattern = parseFloat(value);
        if (!isNaN(fillPattern)) style.FillPattern = fillPattern;
        break;
      case 'Color': // Text color
        const textColor = parseColorEnhanced(value);
        if (textColor) style.TextColor = textColor;
        break;
      case 'TextBkgnd':
        const textBkgnd = parseColorEnhanced(value);
        if (textBkgnd) style.TextBkgnd = textBkgnd;
        break;
      case 'HideText':
        style.HideText = value;
        break;
      case 'TextDirection':
        style.TextDirection = value;
        break;
    }
  });
  
  return style;
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
    
    // Enhanced: Process cell data into proper Style object instead of raw cells
    sheet.Style = processStyleCells(stylesheetObjects[i]['Cell'] || []);

    styleSheets.push(sheet);
  }

  return styleSheets;
};

const parseMastersFile = (jsonObj: any) => {
  const mastersList: VisioMaster[] = [];
  const masterObjects = jsonObj['Masters']['Master'];

  for (let i = 0; i < masterObjects.length; i++) {
    const master = {} as VisioMaster;
    master.Id = masterObjects[i]['$']['ID'];
    master.Name = masterObjects[i]['$']['Name'];
    master.UniqueID = masterObjects[i]['$']['UniqueID'];
    master.BaseID = masterObjects[i]['$']['BaseID'];
    master.MasterType = masterObjects[i]['$']['MasterType'];
    master.RelationshipId = masterObjects[i]['Rel'][0]['$']['r:id'];
    master.Hidden = masterObjects[i]['$']['Hidden'];
    
    // Enhanced: Extract style references from PageSheet
    if (masterObjects[i]['PageSheet']) {
      master.LineStyleRefId = masterObjects[i]['PageSheet']['LineStyle'];
      master.FillStyleRefId = masterObjects[i]['PageSheet']['FillStyle'];
      master.TextStyleRefId = masterObjects[i]['PageSheet']['TextStyle'];
    }

    mastersList.push(master);
  }

  return mastersList;
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
      const cells = shapeContainer['Cell'] || [];

      shape.Id = shapeContainer['$']['ID'];
      shape.MasterId = shapeContainer['$']['Master'];
      
      // Enhanced: Extract style references from shape
      const styleRefs = extractStyleReferences(cells);
      shape.LineStyleRefId = styleRefs.LineStyleRefId || '';
      shape.FillStyleRefId = styleRefs.FillStyleRefId || '';
      shape.TextStyleRefId = styleRefs.TextStyleRefId || '';
      
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

      // Enhanced: Resolve style using inheritance chain (Shape → Master → Stylesheet)
      shape.Style = resolveShapeStyleWithInheritance(shapeContainer, master);

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

const parseLineWeight = (value: string, unit?: string): number => {
  const numValue = parseFloat(value);
  if (isNaN(numValue)) return 0.75; // Default line weight
  
  // Convert to pixels based on unit
  switch (unit) {
    case 'PT': return numValue * 96 / 72; // Points to pixels
    case 'IN': return numValue * 96;      // Inches to pixels
    case 'MM': return numValue * 96 / 25.4; // Millimeters to pixels
    case 'CM': return numValue * 96 / 2.54; // Centimeters to pixels
    default: return numValue * 96; // Assume inches if no unit specified
  }
};

const parseColorEnhanced = (value: string): string => {
  if (!value || value === '' || value === '0') return '';
  
  // Already a hex color
  if (value.startsWith('#')) return value;
  
  // Handle RGB format (e.g., "255,0,0")
  if (value.includes(',')) {
    const parts = value.split(',').map(p => parseInt(p.trim()));
    if (parts.length === 3 && parts.every(p => p >= 0 && p <= 255)) {
      return `#${parts.map(p => p.toString(16).padStart(2, '0')).join('')}`;
    }
  }
  
  // Handle Visio color index
  const colorIndex = parseInt(value);
  if (!isNaN(colorIndex)) {
    return getVisioColorByIndex(colorIndex);
  }
  
  return value;
};

const getVisioColorByIndex = (index: number): string => {
  // Extended Visio color palette
  const visioColors: { [key: number]: string } = {
    0: '#FFFFFF', // White
    1: '#000000', // Black
    2: '#FF0000', // Red
    3: '#00FF00', // Green
    4: '#0000FF', // Blue
    5: '#FFFF00', // Yellow
    6: '#FF00FF', // Magenta
    7: '#00FFFF', // Cyan
    8: '#800000', // Dark Red
    9: '#008000', // Dark Green
    10: '#000080', // Dark Blue
    11: '#808000', // Olive
    12: '#800080', // Purple
    13: '#008080', // Teal
    14: '#C0C0C0', // Silver
    15: '#808080', // Gray
    16: '#9999FF', // Light Blue
    17: '#993366', // Dark Pink
    18: '#FFFFCC', // Light Yellow
    19: '#CCFFFF', // Light Cyan
    20: '#660066', // Dark Purple
    21: '#FF8080', // Light Red
    22: '#0066CC', // Medium Blue
    23: '#CCCCFF', // Very Light Blue
    24: '#000080', // Navy
    25: '#FF00FF', // Fuchsia
    // Add more colors as needed
  };
  
  return visioColors[index] || '#000000';
};

const extractStyleReferences = (cells: any[]): any => {
  const refs: any = {};
  
  if (!cells || !Array.isArray(cells)) return refs;
  
  cells.forEach(cell => {
    const name = cell.$?.N;
    const value = cell.$?.V;
    
    if (name === 'LineStyle') refs.LineStyleRefId = value;
    if (name === 'FillStyle') refs.FillStyleRefId = value;
    if (name === 'TextStyle') refs.TextStyleRefId = value;
  });
  
  return refs;
};

const resolveShapeStyleWithInheritance = (shapeContainer: any, master?: VisioMaster): Style => {
  const style = {} as Style;
  
  // 1. Apply defaults first
  applyDefaultStyles(style);
  
  // 2. Apply stylesheet styles if master references them
  if (master) {
    applyStylesheetStyles(style, master);
  }
  
  // 3. Apply master styles (medium priority)
  if (master?.Style) {
    mergeStyles(style, master.Style);
  }
  
  // 4. Apply shape-specific styles (highest priority)
  const shapeStyle = processStyleCells(shapeContainer.Cell || []);
  mergeStyles(style, shapeStyle);
  
  return style;
};

const applyDefaultStyles = (style: Style): void => {
  style.LineWeight = 0.75;
  style.LineColor = '#000000';
  style.LinePattern = 1;
  style.FillPattern = 1;
  style.FillForeground = '#FFFFFF';
  style.FillBackground = '#000000';
  style.BeginArrow = 0;
  style.EndArrow = 0;
  style.BeginArrowSize = 2;
  style.EndArrowSize = 2;
  style.LineCap = 0;
  style.Rounding = 0;
  style.TextColor = '#000000';
  style.TextBkgnd = '#FFFFFF';
  style.HideText = '0';
  style.TextDirection = '0';
};

const applyStylesheetStyles = (style: Style, master: VisioMaster): void => {
  const styleIds = [master.LineStyleRefId, master.FillStyleRefId, master.TextStyleRefId];
  
  styleIds.forEach(styleId => {
    if (styleId) {
      const stylesheet = stylesheets.find(s => s.ID === styleId);
      if (stylesheet?.Style) {
        mergeStyles(style, stylesheet.Style);
      }
    }
  });
};

const mergeStyles = (target: Style, source: Style): void => {
  Object.keys(source).forEach(key => {
    const sourceValue = (source as any)[key];
    if (sourceValue !== undefined && sourceValue !== null && sourceValue !== '') {
      (target as any)[key] = sourceValue;
    }
  });
};

const getStyleFromObject = (cells: any[]): Style => {
  const style = {} as Style;
  
  // Apply defaults first
  style.LineWeight = 0.75;
  style.LineColor = '#000000';
  style.LinePattern = 1;
  style.FillForeground = '#FFFFFF';
  style.FillBackground = '#000000';
  style.BeginArrow = 0;
  style.EndArrow = 0;
  style.BeginArrowSize = 2;
  style.EndArrowSize = 2;
  style.LineCap = 0;
  style.Rounding = 0;
  style.TextColor = '#000000';
  style.FillPattern = 1;
  style.TextBkgnd = '#FFFFFF';
  style.HideText = '0';
  style.TextDirection = '0';
  
  // Then override with actual values from cells
  cells.forEach(cell => {
    const name = cell?.['$']?.['N'];
    const value = cell?.['$']?.['V'];
    const unit = cell?.['$']?.['U'];
    const formula = cell?.['$']?.['F'];
    
    if (!name || value === undefined) return;
    
    // Skip inherited and themed values for now
    if (formula === 'Inh' || value === 'Themed') return;
    
    switch (name) {
      case 'LineWeight':
        style.LineWeight = parseLineWeight(value, unit);
        break;
      case 'LineColor':
        const lineColor = parseColorEnhanced(value);
        if (lineColor) style.LineColor = lineColor;
        break;
      case 'LinePattern':
        const linePattern = parseFloat(value);
        if (!isNaN(linePattern)) style.LinePattern = linePattern;
        break;
      case 'Rounding':
        const rounding = parseFloat(value);
        if (!isNaN(rounding)) style.Rounding = rounding;
        break;
      case 'BeginArrow':
        const beginArrow = parseInt(value);
        if (!isNaN(beginArrow)) style.BeginArrow = beginArrow;
        break;
      case 'BeginArrowSize':
        const beginArrowSize = parseInt(value);
        if (!isNaN(beginArrowSize)) style.BeginArrowSize = beginArrowSize;
        break;
      case 'EndArrow':
        const endArrow = parseInt(value);
        if (!isNaN(endArrow)) style.EndArrow = endArrow;
        break;
      case 'EndArrowSize':
        const endArrowSize = parseInt(value);
        if (!isNaN(endArrowSize)) style.EndArrowSize = endArrowSize;
        break;
      case 'LineCap':
        const lineCap = parseInt(value);
        if (!isNaN(lineCap)) style.LineCap = lineCap;
        break;
      case 'FillForegnd':
        const fillForeground = parseColorEnhanced(value);
        if (fillForeground) style.FillForeground = fillForeground;
        break;
      case 'FillBkgnd':
        const fillBackground = parseColorEnhanced(value);
        if (fillBackground) style.FillBackground = fillBackground;
        break;
      case 'FillPattern':
        const fillPattern = parseFloat(value);
        if (!isNaN(fillPattern)) style.FillPattern = fillPattern;
        break;
      case 'Color': // Text color
        const textColor = parseColorEnhanced(value);
        if (textColor) style.TextColor = textColor;
        break;
      case 'TextBkgnd':
        const textBkgnd = parseColorEnhanced(value);
        if (textBkgnd) style.TextBkgnd = textBkgnd;
        break;
      case 'HideText':
        style.HideText = value;
        break;
      case 'TextDirection':
        style.TextDirection = value;
        break;
    }
  });

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
