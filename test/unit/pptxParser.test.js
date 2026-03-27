import { describe, it, expect, beforeAll } from 'vitest';
import { readFileSync } from 'node:fs';
import { resolve, dirname } from 'node:path';
import { fileURLToPath } from 'node:url';
import { parsePptx, resolveColors, resolveInheritedStyles } from '../../src/server/converter/pptxParser.js';

const __dirname = dirname(fileURLToPath(import.meta.url));
const fixturePath = resolve(__dirname, '../fixtures/basic.pptx');

describe('pptxParser', () => {
  let parsed;
  let resolved;

  beforeAll(async () => {
    const buffer = readFileSync(fixturePath);
    parsed = await parsePptx(buffer);
    resolved = resolveColors(resolveInheritedStyles(parsed));
  });

  describe('slide discovery', () => {
    it('finds the correct number of slides', () => {
      expect(resolved.slides).toHaveLength(2);
    });

    it('slides are in order', () => {
      expect(resolved.slides[0].index).toBe(1);
      expect(resolved.slides[1].index).toBe(2);
    });
  });

  describe('theme parsing', () => {
    it('extracts theme colors', () => {
      expect(resolved.theme.colors.accent1).toBe('4472C4');
      expect(resolved.theme.colors.dk1).toBe('000000');
      expect(resolved.theme.colors.lt1).toBe('FFFFFF');
    });

    it('extracts theme fonts', () => {
      expect(resolved.theme.fonts.major).toBe('Calibri Light');
      expect(resolved.theme.fonts.minor).toBe('Calibri');
    });
  });

  describe('slide 1 — text extraction (verbatim)', () => {
    let slide1;

    beforeAll(() => {
      slide1 = resolved.slides[0];
    });

    it('extracts title text verbatim', () => {
      const title = slide1.elements.find(
        e => e.type === 'shape' && e.placeholder?.type === 'title'
      );
      expect(title).toBeDefined();
      const allText = title.paragraphs
        .flatMap(p => p.runs.filter(r => r.type === 'text').map(r => r.text))
        .join('');
      expect(allText).toBe('Welcome to Pamphlet');
    });

    it('extracts body text verbatim across all runs', () => {
      const body = slide1.elements.find(
        e => e.type === 'shape' && e.placeholder?.type === 'body'
      );
      expect(body).toBeDefined();
      const para1Text = body.paragraphs[0].runs
        .filter(r => r.type === 'text')
        .map(r => r.text)
        .join('');
      expect(para1Text).toBe('This is bold and italic blue text.');
    });

    it('preserves individual run text segments', () => {
      const body = slide1.elements.find(
        e => e.type === 'shape' && e.placeholder?.type === 'body'
      );
      const runs = body.paragraphs[0].runs.filter(r => r.type === 'text');
      expect(runs).toHaveLength(5);
      expect(runs[0].text).toBe('This is ');
      expect(runs[1].text).toBe('bold');
      expect(runs[2].text).toBe(' and ');
      expect(runs[3].text).toBe('italic blue');
      expect(runs[4].text).toBe(' text.');
    });

    it('extracts free text box content', () => {
      const textBox = slide1.elements.find(
        e => e.type === 'shape' && e.name === 'TextBox 3'
      );
      expect(textBox).toBeDefined();
      const text = textBox.paragraphs[0].runs
        .filter(r => r.type === 'text')
        .map(r => r.text)
        .join('');
      expect(text).toBe('Footer note: Confidential');
    });
  });

  describe('slide 1 — formatting extraction', () => {
    let slide1;

    beforeAll(() => {
      slide1 = resolved.slides[0];
    });

    it('title has correct formatting: Arial, 44pt, bold, red', () => {
      const title = slide1.elements.find(
        e => e.type === 'shape' && e.placeholder?.type === 'title'
      );
      const run = title.paragraphs[0].runs.find(r => r.type === 'text');
      expect(run.props.font).toBe('Arial');
      expect(run.props.size).toBe(4400); // hundredths of pt
      expect(run.props.bold).toBe(true);
      expect(run.props.color).toEqual({ type: 'rgb', value: 'FF0000' });
    });

    it('body bold run is bold', () => {
      const body = slide1.elements.find(
        e => e.type === 'shape' && e.placeholder?.type === 'body'
      );
      const boldRun = body.paragraphs[0].runs.find(
        r => r.type === 'text' && r.text === 'bold'
      );
      expect(boldRun.props.bold).toBe(true);
    });

    it('italic blue run has italic and blue color', () => {
      const body = slide1.elements.find(
        e => e.type === 'shape' && e.placeholder?.type === 'body'
      );
      const italicRun = body.paragraphs[0].runs.find(
        r => r.type === 'text' && r.text === 'italic blue'
      );
      expect(italicRun.props.italic).toBe(true);
      expect(italicRun.props.color).toEqual({ type: 'rgb', value: '0000FF' });
    });

    it('underlined run has underline and scheme color resolved to RGB', () => {
      const body = slide1.elements.find(
        e => e.type === 'shape' && e.placeholder?.type === 'body'
      );
      const underlinedRun = body.paragraphs[1].runs.find(r => r.type === 'text');
      expect(underlinedRun.props.underline).toBe('sng');
      expect(underlinedRun.props.font).toBe('Georgia');
      expect(underlinedRun.props.size).toBe(2000);
      // accent1 → 4472C4 via theme
      expect(underlinedRun.props.color).toEqual({ type: 'rgb', value: '4472C4' });
    });

    it('footer text box has correct small font', () => {
      const textBox = slide1.elements.find(
        e => e.type === 'shape' && e.name === 'TextBox 3'
      );
      const run = textBox.paragraphs[0].runs.find(r => r.type === 'text');
      expect(run.props.font).toBe('Courier New');
      expect(run.props.size).toBe(1400);
      expect(run.props.color).toEqual({ type: 'rgb', value: '888888' });
    });
  });

  describe('slide 2 — group shapes', () => {
    let slide2;

    beforeAll(() => {
      slide2 = resolved.slides[1];
    });

    it('finds group shape with nested shapes', () => {
      const group = slide2.elements.find(e => e.type === 'group');
      expect(group).toBeDefined();
      expect(group.elements).toHaveLength(2);
    });

    it('extracts text from grouped shapes verbatim', () => {
      const group = slide2.elements.find(e => e.type === 'group');
      const textA = group.elements[0].paragraphs[0].runs
        .filter(r => r.type === 'text')
        .map(r => r.text)
        .join('');
      expect(textA).toBe('Grouped text A');

      const textB = group.elements[1].paragraphs[0].runs
        .filter(r => r.type === 'text')
        .map(r => r.text)
        .join('');
      expect(textB).toBe('Grouped text B');
    });

    it('grouped shape B has bold green formatting', () => {
      const group = slide2.elements.find(e => e.type === 'group');
      const run = group.elements[1].paragraphs[0].runs.find(r => r.type === 'text');
      expect(run.props.bold).toBe(true);
      expect(run.props.color).toEqual({ type: 'rgb', value: '009900' });
      expect(run.props.font).toBe('Verdana');
    });
  });

  describe('slide 2 — unsupported objects', () => {
    it('detects SmartArt as unsupported', () => {
      expect(resolved.unsupportedObjects).toContainEqual(
        expect.objectContaining({
          slide: 2,
          description: expect.stringContaining('SmartArt'),
        })
      );
    });

    it('adds unsupported element to slide elements', () => {
      const slide2 = resolved.slides[1];
      const unsupported = slide2.elements.find(e => e.type === 'unsupported');
      expect(unsupported).toBeDefined();
      expect(unsupported.description).toContain('SmartArt');
    });
  });

  describe('master styles and color map', () => {
    it('parses master color map', () => {
      expect(resolved.masterStyles.colorMap).toBeDefined();
      expect(resolved.masterStyles.colorMap.tx1).toBe('dk1');
      expect(resolved.masterStyles.colorMap.bg1).toBe('lt1');
    });

    it('parses master title style defaults', () => {
      expect(resolved.masterStyles.title[1]).toBeDefined();
      expect(resolved.masterStyles.title[1].size).toBe(4400);
    });

    it('parses master body style defaults', () => {
      expect(resolved.masterStyles.body[1]).toBeDefined();
      expect(resolved.masterStyles.body[1].size).toBe(2400);
    });
  });
});
