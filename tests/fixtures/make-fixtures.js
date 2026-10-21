// Generates the small, fictional CSV fixtures used by smoke.spec.js and window.__demo.
// Run with: node tests/fixtures/make-fixtures.js
const fs = require('fs');
const path = require('path');
const out = (name, data) => fs.writeFileSync(path.join(__dirname, name), data);

const sales = [
  'order_id,zip,customer,item,qty,total',
  '1001,02134,Zoë Marchetti,Sourdough loaf,2,12.50',
  '1002,01810,Theo Brandt,Maple pecan scone,6,19.80',
  '1003,90210,Ines Calloway,"Cinnamon roll, large",3,11.25',
  '1004,00501,Ravi Lindqvist,抹茶スコーン,4,14.00',
  '1005,10001,Mara Oyelaran,Rye baguette,1,4.75',
  '1006,02139,Zoë Marchetti,Lemon tart,2,9.90',
  '1007,60614,Quinn Halvorsen,Croissant box,12,30.00',
].join('\r\n') + '\r\n';
out('sample-sales.csv', '﻿' + sales);

const routes = [
  'route;driver;stops;start_zip;window',
  'North loop;Pat Ferreira;7;02134;06:00-09:00',
  'Harbor run;Lou Nakamura;5;02110;07:30-10:00',
  'Hillside;Sky Abernathy;9;02445;08:00-11:30',
  'Market street;Robin Achebe;4;02108;09:00-10:30',
  'Airport express;Dana Kowalczyk;3;02128;05:30-07:00',
].join('\n') + '\n';
out('delivery-routes.csv', routes);

// Windows-1252 encoded export (ë = 0xEB, è = 0xE8, û = 0xFB)
const legacy = 'customer;item\nZoë Marchetti;Crème brûlée\n';
out('sales-1252.csv', Buffer.from([...legacy].map((ch) => ({ 'ë': 0xeb, 'è': 0xe8, 'û': 0xfb }[ch] ?? ch.charCodeAt(0)))));

out('broken-quote.csv', 'order_id,item,qty\n1001,Rye loaf,2\n1002,"Seeded bun,3\n1003,Bagel,5\n');
out('empty.csv', '');
