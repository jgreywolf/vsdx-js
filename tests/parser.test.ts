import { fail } from 'assert';
import { describe, expect, it } from 'vitest';
import { VisioFile } from '../src/types';
import { parseVisioFile } from '../src';
import * as fs from 'fs';

const testFilePath = 'tests/Connectors.vsdx';

describe('given a valid Visio file', () => {
  it('should return stylesheets from document properties', async () => {
    const visioFile = (await parseVisioFile(testFilePath)) as VisioFile;

    expect(visioFile.Stylesheets.length).toBe(8);
  });

  it('should parse masters file', async () => {
    const visioFile = (await parseVisioFile(testFilePath)) as VisioFile;

    expect(visioFile.Masters.length).toBe(34);
  });

  it('should parse file from Buffer', async () => {
    const buffer = fs.readFileSync(testFilePath);
    const visioFile = await parseVisioFile(buffer) as VisioFile;

    expect(visioFile.Stylesheets.length).toBe(8);
    expect(visioFile.Masters.length).toBe(34);
  });

  it('should parse shapes and connectors', async () => {
    const visioFile = (await parseVisioFile(testFilePath)) as VisioFile;
    const firstPage = visioFile.Pages[0];
    const shapes = firstPage.Shapes;

    expect(shapes.filter((shape) => !shape.IsEdge).length).toBe(2);
    expect(shapes.filter((shape) => shape.IsEdge).length).toBe(1);
  });

  it('should parse all shapes correctly', async () => {
    const visioFile = (await parseVisioFile('tests/BasicShapes.vsdx')) as VisioFile;
    const firstPage = visioFile.Pages[0];
    const shapes = firstPage.Shapes;

    expect(shapes[0].Type).toBe('Rectangle');
    expect(shapes[1].Type).toBe('Square');
    expect(shapes[2].Type).toBe('Circle');
    expect(shapes[3].Type).toBe('Triangle');
    expect(shapes[4].Type).toBe('Hexagon');
    expect(shapes[5].Type).toBe('Can');
    expect(shapes[6].Type).toBe('Parallelogram');
    expect(shapes[7].Type).toBe('Trapezoid');
    expect(shapes[8].Type).toBe('Diamond');
    expect(shapes[9].Type).toBe('Rounded Rectangle');
    expect(shapes[10].Type).toBe('Center Drag Circle');
    expect(shapes[11].Type).toBe('Left Brace');
    expect(shapes[12].Type).toBe('Right Brace');
  });
});
