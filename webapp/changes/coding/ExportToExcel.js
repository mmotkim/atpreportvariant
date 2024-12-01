sap.ui.define(
  [
    "sap/ui/core/mvc/ControllerExtension",
    // ,'sap/ui/core/mvc/OverrideExecution'
    "exceljs",
    "pdf-lib",
  ],
  function (
    ControllerExtension,
    // ,OverrideExecution
    ExcelJS,
    PDFLib
  ) {
    "use strict";
    return ControllerExtension.extend("customer.atpreportvariant.ExportToExcel", {
      onClick: async function () {
        var data = this.getColumnsAndRows();
        //initialize worksheet
        const workbook = new ExcelJS.Workbook();
        const worksheet = workbook.addWorksheet("My Sheet", {
          pageSetup: { paperSize: 9, orientation: "landscape" },
        });

        //initial formatting
        worksheet.mergeCells("C1", "Q7");
        worksheet.getCell("C1").value = "Available-to-Promise Report";

        //add all those data to the actual worksheet
        worksheet.columns = data.columns;
        worksheet.addRows(data.rows);
        const row = worksheet.lastrow;

        //download the file
        let buffer = await workbook.xlsx.writeBuffer();

        let blob2 = new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
        let link2 = document.createElement("a");
        link2.href = URL.createObjectURL(blob2);
        link2.download = "fName.xlsx";
        link2.click();
        URL.revokeObjectURL(link2.href);
      },

      onClickPdf: async function () {
        var data = this.getColumnsAndRows();

        console.log("onClickPdf");
        const pdfDoc = await PDFLib.PDFDocument.create();
        const page = pdfDoc.addPage([550, 750]);
        const pdfBytes = await pdfDoc.save();
        // Create a Blob from the PDF bytes
        const blob = new Blob([pdfBytes], { type: "application/pdf" });

        // Create a link element
        const link = document.createElement("a");
        link.href = URL.createObjectURL(blob);
        link.download = "document.pdf";

        // Append the link to the body
        document.body.appendChild(link);

        // Trigger the download
        link.click();

        // Remove the link from the document
        document.body.removeChild(link);
        console.log(pdfBytes);
      },

      getColumnsAndRows: function () {
        var data = {
          columns: [],
          rows: [],
        };
        const tableId = "com.atpreport.materialstock2::sap.suite.ui.generic.template.AnalyticalListPage.view.AnalyticalListPage::Z_QUERY_MATERIALSTOCK--table";
        var oModel = this.getView().getModel();
        var oTable = this.getView().byId(tableId);
        console.log(this.getView().getModel().getProperty("/"));
        console.log("oModel");
        console.log(oModel);
        console.log("oModel.oData");
        console.log(oModel.oData);
        console.log("oTable.getTable().getColumns()");
        console.log(oTable.getTable().getColumns());

        //initialize the visible columns array
        var visibleColumns = [];
        for (var i = 0; i < oTable.getTable().getColumns().length; i++) {
          if (oTable.getTable().getColumns()[i].getProperty("visible")) {
            var column = oTable.getTable().getColumns()[i];
            var headerColId = column.getAggregation("header").getId();
            var headerColObj = sap.ui.getCore().byId(headerColId);
            var headerColValue = headerColObj.getText();
            var key = column.getAggregation("customData")[0].getValue().columnKey;

            visibleColumns.push({ name: headerColValue, id: key });
          }
        }
        console.log("visibleColumns");
        console.log(visibleColumns);

        const worksheetColumns = [];

        visibleColumns.forEach((column) => {
          worksheetColumns.push({ header: column.name, key: column.id, width: 15 });
        });

        //initialize the rows array
        //I fucking give up and just iterate from UI (uses oModel.oData)
        var allRows = oModel.oData;
        var filteredData = Object.values(allRows).filter((item) => item.Matnr !== undefined && item.Matnr !== null);
        console.log(filteredData);

        var rowsToAdd = [];
        filteredData.forEach((dataRow) => {
          const row = visibleColumns.map((col) => dataRow[col.id] || null); // If key is missing, insert null
          rowsToAdd.push(row);
        });
        console.log("rowsToAdd");
        console.log(rowsToAdd);

        data.columns = worksheetColumns;
        data.rows = rowsToAdd;
        return data;
      },

      // metadata: {
      // 	// extension can declare the public methods
      // 	// in general methods that start with "_" are private
      // 	methods: {
      // 		publicMethod: {
      // 			public: true /*default*/ ,
      // 			final: false /*def`ault*/ ,
      // 			overrideExecution: OverrideExecution.Instead /*default*/
      // 		},
      // 		finalPublicMethod: {
      // 			final: true
      // 		},
      // 		onMyHook: {
      // 			public: true /*default*/ ,
      // 			final: false /*default*/ ,
      // 			overrideExecution: OverrideExecution.After
      // 		},
      // 		couldBePrivate: {
      // 			public: false
      // 		}
      // 	}
      // },
      // // adding a private method, only accessible from this controller extension
      // _privateMethod: function() {},
      // // adding a public method, might be called from or overridden by other controller extensions as well
      // publicMethod: function() {},
      // // adding final public method, might be called from, but not overridden by other controller extensions as well
      // finalPublicMethod: function() {},
      // // adding a hook method, might be called by or overridden from other controller extensions
      // // override these method does not replace the implementation, but executes after the original method
      // onMyHook: function() {},
      // // method public per default, but made private via metadata
      // couldBePrivate: function() {},
      // // this section allows to extend lifecycle hooks or override public methods of the base controller
      // override: {
      // 	/**
      // 	 * Called when a controller is instantiated and its View controls (if available) are already created.
      // 	 * Can be used to modify the View before it is displayed, to bind event handlers and do other one-time initialization.
      // 	 * @memberOf {{controllerExtPath}}
      // 	 */
      // 	onInit: function() {
      // 	},
      // 	/**
      // 	 * Similar to onAfterRendering, but this hook is invoked before the controller's View is re-rendered
      // 	 * (NOT before the first rendering! onInit() is used for that one!).
      // 	 * @memberOf {{controllerExtPath}}
      // 	 */
      // 	onBeforeRendering: function() {
      // 	},
      // 	/**
      // 	 * Called when the View has been rendered (so its HTML is part of the document). Post-rendering manipulations of the HTML could be done here.
      // 	 * This hook is the same one that SAPUI5 controls get after being rendered.
      // 	 * @memberOf {{controllerExtPath}}
      // 	 */
      // 	onAfterRendering: function() {
      // 	},
      // 	/**
      // 	 * Called when the Controller is destroyed. Use this one to free resources and finalize activities.
      // 	 * @memberOf {{controllerExtPath}}
      // 	 */
      // 	onExit: function() {
      // 	},
      // 	// override public method of the base controller
      // 	basePublicMethod: function() {
      // 	}
      // }
    });
  }
);
