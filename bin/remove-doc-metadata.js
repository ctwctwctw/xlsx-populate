#!/usr/bin/env node
"use strict";

const XlsxPopulate = require('../lib/XlsxPopulate');

const file = process.argv[2];
const password = process.argv[3];

if (!file) {
    console.error('Usage: remove-doc-metadata <file> [password]');
    process.exit(1);
}

const loadOpts = password ? { password } : {};
const saveOpts = { excludeDocMetadata: true };
if (password) saveOpts.password = password;

XlsxPopulate.fromFileAsync(file, loadOpts)
    .then(workbook => workbook.toFileAsync(file, saveOpts))
    .then(() => console.log('Done'))
    .catch(err => {
        console.error(err.message);
        process.exit(1);
    });
