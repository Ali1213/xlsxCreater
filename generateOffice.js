/*jslint nomen: true */
/*globals module, require, __dirname */
var fs = require('fs');
var path = require('path');
var archiver = require('archiver');

var archive = archiver('zip');

module.exports = function (officeFile, filepath,cb) {
    'use strict';

        var output = fs.createWriteStream(filepath);
        var self = this;
        archive.pipe(output);

        archive.bulk([
            {
                expand: true,
                cwd: officeFile.templateDir,
                src: ['**', '**/.rels']
            }
        ]);

        //archive.directory(officeFile.templateDir, false);

        archive.append(officeFile.sheet1, {
            name: 'xl/worksheets/sheet1.xml'
        });

        archive.append(officeFile.sharedStrings, {
            name: 'xl/sharedStrings.xml'
        });

        archive.finalize();


        archive.on('error', function (err) {
            cb(err);
        });

        output.on('close', function() {
            cb(null,{
                "result":"done"
            });
            console.log("done")
        });
};