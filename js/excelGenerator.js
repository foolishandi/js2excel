class excelGenerator {
    constructor(ExcelJS, options) {
        this.defaultSize = options.defaultSize || []
        this.header = options.header
        this.columns = options.columns
        this.data = options.data
        this.fileName = options.fileName || new Date().toDateString()
        this.sheetName = options.sheetName || "sheet1"
        this.merge = options.merge
        this.autoFilter=options.autoFilter
        this.rangeStyle = options.rangeStyle
        this.cellStyle = options.cellStyle
        this.Workbook = new ExcelJS.Workbook()
        this.workSheet = this.Workbook.addWorksheet(this.sheetName);
        this.init()
    }
    init() {
        let ws = this.workSheet
        ws.defaultRowHeight = this.defaultSize[0] || 15
        ws.defaultColWidth = this.defaultSize[1]/10 || 20
        ws.horizontalCentered=true
        ws.verticalCentered=true
        this._columnsRender(this.columns);
        ws.addRows(this.data);
        this.header && this._headerRender(this.header)
        this.merge && this._mergeRender(this.merge);
        this.cellStyle && this._cellStyleRender(this.cellStyle)
        this.rangeStyle && this._rangeStyleRender(this.rangeStyle);
        if(this.autoFilter) ws.autoFilter=this.autoFilter
        return this
    }
    _columnsRender(columns) {
        let ws = this.workSheet
        let columnsData = columns.map((column) => {
            const width = column.width;
            return {
                header: column.title,
                key: column.dataIndex,
                width: isNaN(width) ? 10 : width / 10,
            };
        });
        ws.columns = columnsData
    }
    _headerRender(options) {
        let ws = this.workSheet
        const header = options.title
        let styles = Object.keys(options.style)
        let cols = this.columns.length
        ws.insertRow(1, [header])
        ws.mergeCells(1, 1, 1, cols)
        styles.forEach(st => {
            ws.getCell("A1")[st] = options.style[st]
        })
    }
    _mergeRender(list) {
        let ws = this.workSheet
        list && list.forEach((item) => {
            ws.mergeCells(item)
        })
    }
    _cellStyleRender(list) {
        let ws = this.workSheet
        list.forEach(item => {
            // console.log(item)
            let type= item.type
            let range=item.range
            let style=item.style
            let styles = Object.keys(style)
            console.log(type, range, style)
            styles.forEach(st => {
                if (type === "cell") {
                    ws.getCell(range)[st] = style[st]
                } else if (type === "column") {
                    ws.getColumn(range)[st] = style[st]
                } else if (type === "row") {
                    ws.getRow(range)[st] = style[st]
                }
            })
        })
    }

    _rangeStyleRender(list) {
        let ws = this.workSheet
        list.map(item => {
            let { range, formulae, style,type='expression' ,operator,text} = item
            ws.addConditionalFormatting({
                ref: range,
                rules: [
                    {
                        type,
                        operator,
                        text,
                        formulae,
                        style
                    }
                ]
            })
        })
    }
    destory(){
        let self=this
        self=null
    }
    saveAsExls() {
        let self = this
        self.Workbook.xlsx.writeBuffer().then(function (buffer) {
            saveAs(
                new Blob([buffer], { type: "application/octet-stream" }),
                `${self.fileName}.xlsx`
            );
        });
    }
}