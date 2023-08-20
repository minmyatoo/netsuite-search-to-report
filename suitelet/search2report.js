/**
 * @NApiVersion 2.1
 * @NScriptType Suitelet
 */
define(['N/ui/serverWidget', 'N/search'], function(serverWidget, search) {
  function onRequest(context) {
    if (context.request.method === 'GET') {
      handleGetRequest(context);
    } else {
      handlePostRequest(context);
    }
  }

  function handleGetRequest(context) {
    const form = createForm();


    context.response.writePage(form);
  }

  function createForm() {
    const form = serverWidget.createForm({
      title: 'Master Reports',
      hideNavBar: false,
    });

    form.addSubmitButton({ label: 'Run Report' });

    const mainGroup = form.addFieldGroup({
      id: 'custpage_main',
      label: 'Main Information',
    });

    const fromDate = form.addField({
      id: 'custpage_fdate',
      type: 'date',
      label: 'From Date',
      container: 'custpage_main',
    });
    fromDate.defaultValue = '1/1/2021';

    const toDate = form.addField({
      id: 'custpage_tdate',
      type: 'date',
      label: 'To Date',
      container: 'custpage_main',
    });
    toDate.defaultValue = new Date();

    const selectReports = form.addField({
      id: 'custpage_reportid',
      label: 'Select Reports',
      type: serverWidget.FieldType.SELECT,
      container: 'custpage_main',
    });
    selectReports.addSelectOption({
      text: 'Production Summary Report',
      value: 'customsearch_prod_summary_cn',//YOUR SAVED SEARCH HERE
    });
    // Add more saved searches


    return form;
  }

  function handlePostRequest(context) {
    const form = createForm();

    form.addButton({
      id: 'custpage_export',
      label: 'Export',
      functionName: 'exportToExcel()',
    });
    const htmlTableField = form.addField({
      id: 'custpage_htmltable',
      label: 'HTML Table',
      type: serverWidget.FieldType.INLINEHTML,
    });

    const fdate_param = context.request.parameters.custpage_fdate;
    const tdate_param = context.request.parameters.custpage_tdate;
    const reportId = context.request.parameters.custpage_reportid;

    const savedSearch = search.load({
      id: reportId,
    });

    savedSearch.filters.push(
      search.createFilter({
        name: 'trandate',
        operator: search.Operator.BETWEEN,
        values: [fdate_param, tdate_param],
      })
    );

    const searchColumns = savedSearch.columns;
    const tableHeader = generateTableHeader(searchColumns);

    const searchResults = runSavedSearch(savedSearch, searchColumns);
    const tableRows = generateTableRows(searchResults, searchColumns);

    const reportName = getReportName(reportId);

    const htmlTable = `
      <div style='background-color: #f2f2f2; padding: 7px;'>
        <h2 style='margin: 0;'>${reportName}</h2>
        <p style='margin: 0;'>From: ${fdate_param} To: ${tdate_param}</p>
      </div>
      <hr>
      <div align='center'>
        <table class='mdl-data-table mdl-js-data-table mdl-data-table--selectable mdl-shadow--2dp'>
          <thead>${tableHeader}</thead>
          <tbody>${tableRows}</tbody>
        </table>
      </div>
    `;

    const sheetJsCdn = `<script src='https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.17.0/xlsx.full.min.js'></script>`;
    const exportToExcelFunction = generateExportToExcelFunction();

    htmlTableField.defaultValue = generateHtmlContent(htmlTable, sheetJsCdn, exportToExcelFunction);

    context.response.writePage(form);
  }

  function generateTableHeader(searchColumns) {
    let tableHeader = '<tr>';
    for (const element of searchColumns) {
      tableHeader += '<th class="mdl-data-table__cell--non-numeric">' + (element.label || element.name) + '</th>';
    }
    tableHeader += '</tr>';
    return tableHeader;
  }

  function runSavedSearch(savedSearch, searchColumns) {
    const searchResults = [];
    savedSearch.run().each(function (result) {
      const row = {};
      for (const element of searchColumns) {
        const column = element;
        const value = result.getValue(column);
        const text = result.getText(column);

        row[column.name] = text && text !== value ? text : value;
      }
      searchResults.push(row);
      return true;
    });
    return searchResults;
  }

  function generateTableRows(searchResults, searchColumns) {
    let tableRows = '';
    for (const resultRow of searchResults) {
      tableRows += '<tr>';
      for (const column of searchColumns) {
        const value = resultRow[column.name] || '';
        tableRows += '<td class="mdl-data-table__cell--non-numeric">' + value + '</td>';
      }
      tableRows += '</tr>';
    }
    return tableRows;
  }

  function getReportName(reportId) {
    const reportObj = {
      customsearch_hc_prod_summary_cn: 'Production Summary Report',
      customsearch_prod_schedule: 'Production Schedule Report by PO reference',
    };
    return reportObj[reportId] || '';
  }

  function generateExportToExcelFunction() {
    return `
      <script>
        function exportToExcel() {
          const table = document.querySelector('.mdl-data-table');
          const wb = XLSX.utils.table_to_book(table, { sheet: 'Sheet1' });
          XLSX.writeFile(wb, 'MasterReports.xlsx');
        }
      </script>
    `;
  }

  function generateHtmlContent(htmlTable, sheetJsCdn, exportToExcelFunction) {
    const mdlCss = `<link rel='stylesheet' href='https://fonts.googleapis.com/icon?family=Material+Icons'>
                    <link rel='stylesheet' href='https://code.getmdl.io/1.3.0/material.indigo-pink.min.css'>`;

    const customCss = `
      <style>
        .report-table {
          border-collapse: collapse;
          width: 100%;
        }
        .report-table th,
        .report-table td {
          padding: 8px;
          text-align: left;
          border-bottom: 1px solid #ddd;
        }
        .report-table th {
          background-color: #f2f2f2;
          font-weight: bold;
          color: #424242;
        }
        .report-table tr:hover {
          background-color: #f5f5f5;
        }
      </style>
    `;

    const mdlJs = `<script defer src='https://code.getmdl.io/1.3.0/material.min.js'></script>`;

    return mdlCss + customCss + htmlTable + mdlJs + sheetJsCdn + exportToExcelFunction;
  }

  return {
    onRequest: onRequest,
  };
});
