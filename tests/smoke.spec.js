// smoke.spec.js — CSV to Excel Converter: core actions on the fictional fixtures in tests/fixtures/
// (regenerate them with `node tests/fixtures/make-fixtures.js`).
const { test, expect } = require('@playwright/test');
const path = require('path');
const fs = require('fs');

const fixture = (name) => path.join(__dirname, 'fixtures', name);

// Reads a downloaded workbook back with the page's own vendored SheetJS copy.
async function readWorkbook(page, file) {
  const b64 = fs.readFileSync(file).toString('base64');
  return page.evaluate((data) => {
    const wb = XLSX.read(data, { type: 'base64' });
    const cfb = XLSX.CFB.read(data, { type: 'base64' });
    const xml = (n) => new TextDecoder().decode(XLSX.CFB.find(cfb, n).content);
    return {
      names: wb.SheetNames,
      sheets: Object.fromEntries(wb.SheetNames.map((n) => [n, XLSX.utils.sheet_to_json(wb.Sheets[n], { header: 1, raw: true })])),
      cells: Object.fromEntries(wb.SheetNames.map((n) => [n, wb.Sheets[n]])),
      sheet1: xml('/xl/worksheets/sheet1.xml'),
      styles: xml('/xl/styles.xml'),
    };
  }, b64);
}

async function download(page) {
  await expect(page.locator('#download')).toBeEnabled();
  const [dl] = await Promise.all([page.waitForEvent('download'), page.click('#download')]);
  const file = await dl.path();
  return { dl, file, size: fs.statSync(file).size };
}

test.describe('CSV to Excel Converter', () => {
  test.beforeEach(async ({ page }) => {
    await page.goto('./');
  });

  test('smoke: dropping sample-sales.csv shows a preview with the header row detected', async ({ page }) => {
    await page.setInputFiles('#file-input', fixture('sample-sales.csv'));
    const card = page.locator('.file-card');
    await expect(card).toHaveCount(1);
    await expect(card.locator('.file-card__meta')).toContainText('comma detected');
    await expect(card.locator('.file-card__meta')).toContainText('8 rows · 6 columns');
    await expect(card.locator('thead th')).toHaveText(['order_id', 'zip', 'customer', 'item', 'qty', 'total']);
    await expect(card.locator('tbody tr')).toHaveCount(5);
    await expect(card.locator('tbody tr').first()).toContainText('Zoë Marchetti');
    await expect(card.locator('tbody')).toContainText('抹茶スコーン');
    await expect(page.locator('#download')).toHaveText('Download sample-sales.xlsx');
    // header off: the first row becomes data
    await page.uncheck('#opt-header');
    await expect(card.locator('thead')).toHaveCount(0);
    await expect(card.locator('tbody tr').first()).toContainText('order_id');
  });

  test('smoke: a semicolon separated sample is detected automatically and split into columns', async ({ page }) => {
    await page.setInputFiles('#file-input', fixture('delivery-routes.csv'));
    const card = page.locator('.file-card');
    await expect(card.locator('.file-card__meta')).toContainText('semicolon detected');
    await expect(card.locator('.file-card__meta')).toContainText('5 columns');
    await expect(card.locator('thead th')).toHaveText(['route', 'driver', 'stops', 'start_zip', 'window']);
    await expect(card.locator('tbody tr').first().locator('td')).toHaveCount(5);
    // manual override to comma collapses it to one column
    await card.locator('select').selectOption(',');
    await expect(card.locator('.file-card__meta')).toContainText('comma (manual) · 6 rows · 1 columns');
    await card.locator('select').selectOption('auto');
    await expect(card.locator('.file-card__meta')).toContainText('semicolon detected');
  });

  test('smoke: download produces a .xlsx file larger than 1 KB with bold, frozen header', async ({ page }) => {
    await page.setInputFiles('#file-input', fixture('sample-sales.csv'));
    const { dl, file, size } = await download(page);
    expect(dl.suggestedFilename()).toBe('sample-sales.xlsx');
    expect(size).toBeGreaterThan(1024);
    const head = fs.readFileSync(file).subarray(0, 2).toString('latin1');
    expect(head).toBe('PK');
    const wb = await readWorkbook(page, file);
    expect(wb.names).toEqual(['sample-sales']);
    expect(wb.sheets['sample-sales'][0]).toEqual(['order_id', 'zip', 'customer', 'item', 'qty', 'total']);
    expect(wb.sheets['sample-sales'][4][3]).toBe('抹茶スコーン');
    expect(wb.cells['sample-sales'].E2).toMatchObject({ t: 'n', v: 2 });
    expect(wb.cells['sample-sales'].F2).toMatchObject({ t: 'n', v: 12.5 });
    expect(wb.sheet1).toContain('state="frozen"');
    expect(wb.styles).toContain('<b/>');
    await expect(page.locator('#results .convert-result')).toContainText('sample-sales.xlsx');
  });

  test('smoke: dropping two CSV files creates a workbook with two sheets, in card order', async ({ page }) => {
    await page.setInputFiles('#file-input', [fixture('sample-sales.csv'), fixture('delivery-routes.csv')]);
    await expect(page.locator('.file-card')).toHaveCount(2);
    await expect(page.locator('#download')).toHaveText('Download workbook.xlsx');
    // reorder with the keyboard-accessible buttons
    await page.getByRole('button', { name: 'Move delivery-routes.csv up' }).click();
    await expect(page.locator('.file-card').first()).toHaveAttribute('data-name', 'delivery-routes.csv');
    const { dl, file } = await download(page);
    expect(dl.suggestedFilename()).toBe('workbook.xlsx');
    const wb = await readWorkbook(page, file);
    expect(wb.names).toEqual(['delivery-routes', 'sample-sales']);
    expect(wb.sheets['delivery-routes'][1]).toEqual(['North loop', 'Pat Ferreira', 7, '02134', '06:00-09:00']);
    // removing a card leaves one sheet
    await page.getByRole('button', { name: 'Remove delivery-routes.csv' }).click();
    await expect(page.locator('.file-card')).toHaveCount(1);
    const second = await download(page);
    expect((await readWorkbook(page, second.file)).names).toEqual(['sample-sales']);
  });

  test('smoke: leading zeros in a zip code column are preserved when text columns are kept as text', async ({ page }) => {
    await page.setInputFiles('#file-input', fixture('sample-sales.csv'));
    await expect(page.locator('#opt-text')).toBeChecked();
    let wb = await readWorkbook(page, (await download(page)).file);
    const zips = wb.sheets['sample-sales'].slice(1).map((r) => r[1]);
    expect(zips).toEqual(['02134', '01810', '90210', '00501', '10001', '02139', '60614']);
    expect(wb.cells['sample-sales'].B2.t).toBe('s');
    // with the option off, Excel-style number conversion drops the zeros
    await page.uncheck('#opt-text');
    wb = await readWorkbook(page, (await download(page)).file);
    expect(wb.cells['sample-sales'].B2).toMatchObject({ t: 'n', v: 2134 });
  });

  test('smoke: Windows-1252 files are decoded and non-Latin text round trips', async ({ page }) => {
    await page.setInputFiles('#file-input', fixture('sales-1252.csv'));
    const card = page.locator('.file-card');
    await expect(card.locator('.file-card__meta')).toContainText('Windows-1252');
    await expect(card.locator('tbody')).toContainText('Zoë Marchetti');
    const wb = await readWorkbook(page, (await download(page)).file);
    expect(wb.sheets['sales-1252'][1]).toEqual(['Zoë Marchetti', 'Crème brûlée']);
  });

  test('smoke: broken, empty and wrong-type files show a clear error', async ({ page }) => {
    await page.setInputFiles('#file-input', [fixture('broken-quote.csv'), fixture('empty.csv'), fixture('sample.png'), fixture('sample-sales.csv')]);
    const errors = page.locator('#results .alert--error');
    await expect(errors.filter({ hasText: 'broken-quote.csv' })).toContainText('line 3');
    await expect(errors.filter({ hasText: 'empty.csv' })).toContainText('empty');
    await expect(errors.filter({ hasText: 'sample.png' })).toContainText('only CSV');
    await expect(page.locator('.file-card')).toHaveCount(1);
    await expect(page.locator('.file-card')).toHaveAttribute('data-name', 'sample-sales.csv');
  });

  test('smoke: works at 390px wide', async ({ page }) => {
    await page.setViewportSize({ width: 390, height: 844 });
    await page.goto('./');
    await page.setInputFiles('#file-input', [fixture('sample-sales.csv'), fixture('delivery-routes.csv')]);
    await expect(page.locator('.file-card thead th').first()).toBeVisible();
    const { size } = await download(page);
    expect(size).toBeGreaterThan(1024);
    await expect(page.locator('#results .convert-result')).toBeVisible();
    const widths = await page.evaluate(() => ({ doc: document.documentElement.scrollWidth, win: window.innerWidth }));
    expect(widths.doc).toBeLessThanOrEqual(widths.win + 1);
    const btn = await page.locator('#download').boundingBox();
    expect(btn.x + btn.width).toBeLessThanOrEqual(390);
  });

  test('smoke: __demo shows a finished two-sheet result', async ({ page }) => {
    await page.evaluate(() => window.__demo());
    await expect(page.locator('.file-card')).toHaveCount(2);
    await expect(page.locator('#results .convert-result')).toContainText('2 sheets: sample-sales, delivery-routes');
  });
});
