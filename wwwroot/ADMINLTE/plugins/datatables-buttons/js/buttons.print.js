(function (factory) {
    if (typeof define === 'function' && define.amd) {
        // AMD
        define(['jquery', 'datatables.net', 'datatables.net-buttons'], function ($) {
            return factory($, window, document);
        });
    }
    else if (typeof exports === 'object') {
        // CommonJS
        module.exports = function (root, $) {
            if (!root) {
                root = window;
            }

            if (!$ || !$.fn.dataTable) {
                $ = require('datatables.net')(root, $).$;
            }

            if (!$.fn.dataTable.Buttons) {
                require('datatables.net-buttons')(root, $);
            }

            return factory($, root, root.document);
        };
    }
    else {
        // Browser
        factory(jQuery, window, document);
    }
}(function ($, window, document, undefined) {
    'use strict';
    var DataTable = $.fn.dataTable;

    DataTable.ext.buttons.downloadWord = {
        className: 'buttons-download-word',

        text: function (dt) {
            return dt.i18n('buttons.downloadWord', 'Download DOCX');
        },

        action: function (e, dt, button, config) {
            var data = dt.buttons.exportData(
                $.extend({ decodeEntities: false }, config.exportOptions) // XSS protection
            );
            var exportInfo = dt.buttons.exportInfo(config);
            var columnClasses = dt
                .columns(config.exportOptions.columns)
                .flatten()
                .map(function (idx) {
                    return dt.settings()[0].aoColumns[dt.column(idx).index()].sClass;
                })
                .toArray();

            var addRow = function (d, tag) {
                var str = '<tr>';

                for (var i = 0, ien = d.length; i < ien; i++) {
                    // null and undefined aren't useful in the print output
                    var dataOut = d[i] === null || d[i] === undefined ?
                        '' :
                        d[i];
                    var classAttr = columnClasses[i] ?
                        ' class="' + columnClasses[i] + '"' :
                        '';

                    str += '<' + tag + classAttr + '>' + dataOut + '</' + tag + '>';
                }

                return str + '</tr>';
            };

            // Construct a table for downloading as DOCX
            var html = '<table>';

            if (config.header) {
                html += '<thead>' + addRow(data.header, 'th') + '</thead>';
            }

            html += '<tbody>';
            for (var i = 0, ien = data.body.length; i < ien; i++) {
                html += addRow(data.body[i], 'td');
            }
            html += '</tbody>';

            if (config.footer && data.footer) {
                html += '<tfoot>' + addRow(data.footer, 'th') + '</tfoot>';
            }
            html += '</table>';

            // Create a Blob from the HTML table
            var blob = new Blob([html], { type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });

            // Create a temporary link element to trigger the download
            var link = document.createElement('a');
            link.href = window.URL.createObjectURL(blob);
            link.download = 'filename.docx'; // Specify the file name
            link.click();

            // Clean up
            window.URL.revokeObjectURL(link.href);
        },

        title: '*',

        exportOptions: {},

        header: true,

        footer: false,

        customize: null
    };

    return DataTable.Buttons;
}));
