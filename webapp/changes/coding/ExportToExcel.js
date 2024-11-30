sap.ui.define(
  [
    "sap/ui/core/mvc/ControllerExtension",
    // ,'sap/ui/core/mvc/OverrideExecution'
    "exceljs",
    "pdfmake",
  ],
  function (
    ControllerExtension,
    // ,OverrideExecution
    ExcelJS,
    pdfMake
  ) {
    "use strict";
    return ControllerExtension.extend(
      "customer.atpreportvariant.ExportToExcel",
      {
        onClickExcel: async function () {
          const tableId =
            "com.atpreport.materialstock2::sap.suite.ui.generic.template.AnalyticalListPage.view.AnalyticalListPage::Z_QUERY_MATERIALSTOCK--table";
          var oModel = this.getView().getModel();
          var oTable = this.getView().byId(tableId);
          console.log(this.getView().getModel().getProperty("/"));
          console.log("oModel");
          console.log(oModel);
          console.log("oModel.oData");
          console.log(oModel.oData);
          console.log("oTable.getTable().getColumns()");
          console.log(oTable.getTable().getColumns());
          // console.log(oTable);
          // var oModel2 = oTable.getModel();
          // console.log(oTable.getItems());
          // console.log(oTable.getRows());
          // console.log(oTable.getTable().getColumns());
          // console.log(oTable.getItems());
          // console.log("model");
          // console.log(oModel);
          // console.log(oModel.getData());
          // console.log("data2");
          // console.log(oTable.getModel().getData());
          // this.getView().setModel(oModel, "tableModel");
          // var oTable = this.getView().byId(tableId);
          // console.log(oTable.getItems()[0].getBindingContext("id-1732831493737-15"));
          // console.log(this.getView().getModel().getProperty(oTable.getItems()[0].getBindingContext().getPath()));
          // console.log(oTable.getRows());
          // console.log("table");
          // console.log(oTable);
          // console.log(oTable.getTableBindingPath());
          // console.log("getTable");
          // console.log(oTable.getTable());
          // console.log("getTable.mAggregation");
          // console.log(oTable.getTable().mAggregations);
          // console.log(oTable.getTable().mAggregations.items);
          // this.getView().setModel(attModel);

          //initialize the visible columns array
          var visibleColumns = [];
          for (var i = 0; i < oTable.getTable().getColumns().length; i++) {
            if (oTable.getTable().getColumns()[i].getProperty("visible")) {
              var column = oTable.getTable().getColumns()[i];
              var headerColId = column.getAggregation("header").getId();
              var headerColObj = sap.ui.getCore().byId(headerColId);
              var headerColValue = headerColObj.getText();
              var key = column
                .getAggregation("customData")[0]
                .getValue().columnKey;

              visibleColumns.push({ name: headerColValue, id: key });
            }
          }
          console.log("visibleColumns");
          console.log(visibleColumns);

          const worksheetColumns = [];

          visibleColumns.forEach((column) => {
            worksheetColumns.push({
              header: column.name,
              key: column.id,
              width: 15,
            });
          });

          //initialize the rows array
          //I fucking give up and just iterate from UI (uses oModel.oData)
          var allRows = oModel.oData;
          var filteredData = Object.values(allRows).filter(
            (item) => item.Matnr !== undefined && item.Matnr !== null
          );
          console.log(filteredData);

          var rowsToAdd = [];
          filteredData.forEach((dataRow) => {
            const row = visibleColumns.map((col) => dataRow[col.id] || null); // If key is missing, insert null
            rowsToAdd.push(row);
          });
          console.log("rowsToAdd");
          console.log(rowsToAdd);

          //initialize worksheet
          const workbook = new ExcelJS.Workbook();
          const worksheet = workbook.addWorksheet("My Sheet", {
            pageSetup: { paperSize: 9, orientation: "landscape" },
          });

          //initial formatting
          worksheet.mergeCells("C1", "Q7");
          worksheet.getCell("C1").value = "Available-to-Promise Report";

          //add all those data to the actual worksheet
          worksheet.columns = worksheetColumns;
          worksheet.addRows(rowsToAdd);
          const row = worksheet.lastrow;

          //download the file
          let buffer = await workbook.xlsx.writeBuffer();

          let blob = new Blob([buffer], {
            type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
          });
          let link = document.createElement("a");
          link.href = URL.createObjectURL(blob);
          link.download = "fName.xlsx";
          link.click();
          URL.revokeObjectURL(link.href);

          //cleanup
          oModel.refresh();
        },

        onClickPdf: async function () {
          const tableId =
            "com.atpreport.materialstock2::sap.suite.ui.generic.template.AnalyticalListPage.view.AnalyticalListPage::Z_QUERY_MATERIALSTOCK--table";
          var oModel = this.getView().getModel();
          var oTable = this.getView().byId(tableId);
          console.log(this.getView().getModel().getProperty("/"));
          console.log("oModel");
          console.log(oModel);
          console.log("oModel.oData");
          console.log(oModel.oData);
          console.log("oTable.getTable().getColumns()");
          console.log(oTable.getTable().getColumns());
          var visibleColumns = [];
          for (var i = 0; i < oTable.getTable().getColumns().length; i++) {
            if (oTable.getTable().getColumns()[i].getProperty("visible")) {
              var column = oTable.getTable().getColumns()[i];
              var headerColId = column.getAggregation("header").getId();
              var headerColObj = sap.ui.getCore().byId(headerColId);
              var headerColValue = headerColObj.getText();
              var key = column
                .getAggregation("customData")[0]
                .getValue().columnKey;

              visibleColumns.push({ name: headerColValue, id: key });
            }
          }
          console.log("visibleColumns");
          console.log(visibleColumns);

          const worksheetColumns = [];

          visibleColumns.forEach((column) => {
            worksheetColumns.push({
              header: column.name,
              key: column.id,
              width: 35,
            });
          });

          //initialize the rows array
          //I fucking give up and just iterate from UI (uses oModel.oData)
          var allRows = oModel.oData;
          var filteredData = Object.values(allRows).filter(
            (item) =>
              item.Matnr !== undefined &&
              item.Matnr !== null &&
              item.Matkl !== undefined &&
              item.Matkl !== null
          );
          console.log(filteredData);

          var rowsToAdd = [];
          filteredData.forEach((dataRow) => {
            const row = visibleColumns.map((col) => dataRow[col.id] || null); // If key is missing, insert null
            rowsToAdd.push(row);
          });
          console.log("rowsToAdd");
          console.log(rowsToAdd);

          //initialize pdf
          var docDefinition = {
            content: [
              'First paragraph',
              'Another paragraph, this time a little bit longer to make sure, this line will be divided into at least two lines'
            ]
          };
          pdfMake.createPdf(docDefinition).download();




          // //initialize pdf
          // const doc = new jsPDF();

          // doc.setFontSize(18);
          // doc.setFont("helvetica", "bold");
          // doc.text("Available-To-Promise Report", 105, 20, { align: "center" });

          // // Add text on the next lines, aligning to the right of the screen
          // doc.setFontSize(12);
          // doc.setFont("helvetica", "normal");
          // doc.text("Company: Yuki", 200, 30, { align: "right" });
          // doc.text("User:", 200, 40, { align: "right" });
          // doc.text("Fiscal Year: 2024", 200, 50, { align: "right" });
          // doc.text("Date:", 200, 60, { align: "right" });

          // //add all those data to the actual worksheet
          // var headers = createHeaders(worksheetColumns);
          // doc.table(10, 70, rowsToAdd, headers, { autoSize: true });

          // //download the file
          // doc.save("ATP Report.pdf");

          //cleanup
          oModel.refresh();
        },

        createHeaders: function (columnList) {
          var result = [];
          columnList.forEach((column) => {
            result.push({
              name: column.name,
              id: column.id,
              width: column.width,
              align: "center",
              padding: 0,
              prompt: column.name,
            });
          });
          return result;
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
      }
    );
  }
);
