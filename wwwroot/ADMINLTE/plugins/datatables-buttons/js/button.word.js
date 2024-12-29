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

	DataTable.ext.buttons.downloadDocx = {
		className: 'buttons-download-docx',

		text: function (dt) {
			return dt.i18n('buttons.downloadDocx', 'Download DOCX');
		},

		action: function (e, dt, button, config) {
			// Get column data from DataTable
			var columnData = dt.columns().data().toArray();

			// Create a new DOCX file
			var doc = new docx.Document();

			// Add table to DOCX file
			var table = doc.createTable(columnData[0].length, columnData.length);
			for (var i = 0; i < columnData.length; i++) {
				for (var j = 0; j < columnData[i].length; j++) {
					table.getCell(j, i).addContent(new docx.Paragraph(columnData[i][j]));
				}
			}

			// Generate the DOCX file
			docx.Packer.toBlob(doc).then(function (blob) {
				// Save the DOCX file
				saveAs(blob, 'data.docx');
			});
		},

		title: '*',

		exportOptions: {},

		customize: null
	};

	return DataTable.Buttons;
}));
