/**
 * Selection drag lifecycle regressions.
 *
 * The painted-page editor owns selection in an off-screen ProseMirror and
 * routes page gestures through global listeners. These tests simulate the
 * release/cancellation signals WKWebView can deliver when a drag crosses the
 * web-view boundary and assert that later pointer motion cannot keep extending
 * the old selection.
 */

import { test, expect, type Page } from '@playwright/test';
import { EditorPage } from '../helpers/editor-page';

const text = 'alpha bravo charlie delta echo foxtrot golf hotel';

async function selectedText(page: Page): Promise<string> {
  return page.evaluate(() => window.getSelection()?.toString() ?? '');
}

test.describe('Paged editor selection drag lifecycle', () => {
  let editor: EditorPage;

  test.beforeEach(async ({ page }) => {
    editor = new EditorPage(page);
    await editor.goto();
    await editor.waitForReady();
    await editor.newDocument();
    await editor.focus();
    await editor.typeText(text);
  });

  test('a no-button move retires a lost mouseup before extending selection', async ({ page }) => {
    const span = page.locator(`.layout-page span:has-text("${text}")`).first();
    const box = await span.boundingBox();
    expect(box).not.toBeNull();
    if (!box) return;

    const y = box.y + box.height / 2;
    const startX = box.x + 4;
    const selectedToX = box.x + box.width * 0.42;
    const staleMoveX = box.x + box.width * 0.88;

    await page.mouse.move(startX, y);
    await page.mouse.down();
    await page.mouse.move(selectedToX, y, { steps: 4 });
    const selectionAtRelease = await selectedText(page);
    expect(selectionAtRelease.length).toBeGreaterThan(0);

    // Simulate re-entry after the physical button was released outside
    // WKWebView: WebKit reports a move with no primary button, but no mouseup.
    await page.evaluate(
      ({ x, y }) =>
        window.dispatchEvent(
          new MouseEvent('mousemove', { clientX: x, clientY: y, buttons: 0, bubbles: true })
        ),
      { x: staleMoveX, y }
    );

    expect(await selectedText(page)).toBe(selectionAtRelease);

    // Further hover movement must remain inert rather than following the mouse.
    await page.evaluate(
      ({ x, y }) =>
        window.dispatchEvent(
          new MouseEvent('mousemove', { clientX: x, clientY: y, buttons: 0, bubbles: true })
        ),
      { x: box.x + box.width - 2, y }
    );
    expect(await selectedText(page)).toBe(selectionAtRelease);

    // Reset Playwright's own held-button state; the editor gesture is already done.
    await page.mouse.up();
  });

  test('pointer cancellation prevents later primary-button movement extending the old drag', async ({
    page,
  }) => {
    const span = page.locator(`.layout-page span:has-text("${text}")`).first();
    const box = await span.boundingBox();
    expect(box).not.toBeNull();
    if (!box) return;

    const y = box.y + box.height / 2;
    await page.mouse.move(box.x + 4, y);
    await page.mouse.down();
    await page.mouse.move(box.x + box.width * 0.36, y, { steps: 3 });
    const selectionAtCancel = await selectedText(page);
    expect(selectionAtCancel.length).toBeGreaterThan(0);

    await page.evaluate(() => window.dispatchEvent(new PointerEvent('pointercancel')));
    await page.evaluate(
      ({ x, y }) =>
        window.dispatchEvent(
          new MouseEvent('mousemove', { clientX: x, clientY: y, buttons: 1, bubbles: true })
        ),
      { x: box.x + box.width * 0.9, y }
    );

    expect(await selectedText(page)).toBe(selectionAtCancel);
    await page.mouse.up();
  });
});
