/**
 * Issue #811 — an OOXML `srcRect` crop must render by scaling the cropped
 * region to fill the display box (inside an overflow-hidden wrapper), NOT by
 * stretching the whole source into the box and clipping it (which squashes the
 * image, losing its aspect ratio).
 */
import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, beforeAll, describe, expect, test } from 'bun:test';
import { applyImageCrop, hasImageCrop } from '../renderImage';

beforeAll(() => GlobalRegistrator.register());
afterAll(() => GlobalRegistrator.unregister());

describe('applyImageCrop (#811)', () => {
  test('a vertically-cropped image keeps its source aspect ratio', () => {
    // The Format_test.docx case: srcRect t=25.793% b=30.317%, displayed 624×183.
    const img = document.createElement('img');
    const el = applyImageCrop(img, { cropTop: 0.25793, cropBottom: 0.30317 }, 624, 183, document);

    // Returns an overflow-hidden wrapper at the display size.
    expect(el).not.toBe(img);
    expect(el.style.overflow).toBe('hidden');
    expect(el.style.width).toBe('624px');
    expect(el.style.height).toBe('183px');

    // The img is scaled to the FULL (uncropped) height so the visible 43.89%
    // band fills the 183px box: 183 / (1 - .25793 - .30317) ≈ 417px. Width is
    // unchanged (no horizontal crop). 624×417 ≈ the source's own 1.5:1 ratio —
    // i.e. NOT squashed into 624×183 (3.4:1).
    expect(Math.round(parseFloat(img.style.height))).toBe(417);
    expect(img.style.width).toBe('624px');
    // Offset up by the trimmed top fraction so the band is centered.
    expect(Math.round(parseFloat(img.style.marginTop))).toBe(-108);
    expect(img.style.clipPath).toBe('');
  });

  test('an uncropped image is returned untouched (no wrapper)', () => {
    const img = document.createElement('img');
    expect(hasImageCrop({})).toBe(false);
    expect(applyImageCrop(img, {}, 100, 100, document)).toBe(img);
  });
});
