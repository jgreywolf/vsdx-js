export interface VisioFile {
  Masters: VisioMaster[];
  Pages: VisioPage[];
  Relationships: VisioRelationship[];
  Stylesheets: VisioStylesheet[];
  // currently unused
  Settings?: string;
}

export interface VisioEntity {
  Id: string;
  Name: string;
}

export interface Style {
  FillForeground: string;
  FillBackground: string;
  FillPattern: number;
  TextColor: string;
  TextBkgnd: string;
  HideText: string;
  TextDirection: string;
  LineWeight: number;
  LineColor: string;
  LinePattern: number;
  Rounding: number;
  BeginArrow: number;
  BeginArrowSize: number;
  EndArrow: number;
  EndArrowSize: number;
  LineCap: number;  
}

export interface VisioRelationship extends VisioEntity {
  Target: string;
  Type: 'Master' | 'Page';
}

export interface VisioMaster extends VisioEntity {
  UniqueID: string;
  BaseID: string;
  MasterType: string;
  RelationshipId: string;
  Hidden: string;
  // these fields point to the style xml files, and map to RelationshipId in each entity
  LineStyleRefId: string;
  FillStyleRefId: string;
  TextStyleRefId: string;
}

export interface VisioPage extends VisioEntity {
  Shapes: VisioShape[];
  RelationshipId: string;
}

export interface VisioShape extends VisioEntity {
  // maps to VisioMaster.Id
  MasterId: string;
  Type: string;
  Label: string;
  IsEdge: boolean;
  FromNode: string;
  ToNode: string;
  // these fields point to the style xml files, and map to RelationshipId in each entity
  LineStyleRefId: string;
  FillStyleRefId: string;
  TextStyleRefId: string;
  Style: Style;
}

export interface VisioStylesheet {
  ID: string;
  Name: string;
  LineStyleRefId: string;
  FillStyleRefId: string;
  TextStyleRefId: string;
  Style: Style;
}