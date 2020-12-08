/*
 * This software is in the public domain under CC0 1.0 Universal plus a
 * Grant of Patent License.
 *
 * To the extent possible under law, the author(s) have dedicated all
 * copyright and related and neighboring rights to this software to the
 * public domain worldwide. This software is distributed without any
 * warranty.
 *
 * You should have received a copy of the CC0 Public Domain Dedication
 * along with this software (see the LICENSE.md file). If not, see
 * <http://creativecommons.org/publicdomain/zero/1.0/>.
 */
package org.moqui.poi

import com.alibaba.excel.EasyExcel
import com.alibaba.excel.ExcelWriter
import com.alibaba.excel.converters.Converter
import com.alibaba.excel.enums.CellDataTypeEnum
import com.alibaba.excel.metadata.CellData
import com.alibaba.excel.metadata.GlobalConfiguration
import com.alibaba.excel.metadata.Head
import com.alibaba.excel.metadata.property.ExcelContentProperty
import com.alibaba.excel.write.handler.RowWriteHandler
import com.alibaba.excel.write.handler.SheetWriteHandler
import com.alibaba.excel.write.metadata.WriteSheet
import com.alibaba.excel.write.metadata.WriteTable
import com.alibaba.excel.write.metadata.holder.WriteSheetHolder
import com.alibaba.excel.write.metadata.holder.WriteTableHolder
import com.alibaba.excel.write.metadata.holder.WriteWorkbookHolder
import com.alibaba.excel.write.style.AbstractCellStyleStrategy
import com.alibaba.excel.write.style.column.LongestMatchColumnWidthStyleStrategy
import groovy.transform.CompileStatic
import org.apache.poi.ss.usermodel.*
import org.apache.poi.ss.util.DateFormatConverter
import org.moqui.entity.EntityValue
import org.moqui.impl.context.ExecutionContextImpl
import org.moqui.impl.entity.EntityDefinition
import org.moqui.impl.screen.ScreenDefinition
import org.moqui.impl.screen.ScreenForm
import org.moqui.util.MNode
import org.moqui.util.StringUtilities
import org.slf4j.Logger
import org.slf4j.LoggerFactory

/**
 * Generate an Excel (XLSX) file with form-list based output similar to CSV export.
 *
 * Note that this relies on internal classes in Moqui Framework, ie an unstable API.
 */
@CompileStatic
class FormListEasyExcelRender {
    private static final Logger logger = LoggerFactory.getLogger(FormListEasyExcelRender.class)

    // ignore the same field elements as the DefaultScreenMacros.csv.ftl file
    protected static Set<String> ignoreFieldElements = new HashSet<>(["ignored", "hidden", "submit", "image",
                                                                      "date-find", "file", "password", "range-find", "reset"])

    protected ScreenForm.FormInstance formInstance
    protected ExecutionContextImpl eci

    static void renderScreen(ScreenDefinition sd, ExecutionContextImpl eci, OutputStream os) {
        ArrayList<ScreenForm> formList = sd.getAllForms()

        ExcelWriter excelWriter = EasyExcel.write(os).build()

        for (ScreenForm form in formList) {
            MNode formNode = form.getOrCreateFormNode()
            if (formNode.name != "form-list") continue

            ScreenForm.FormInstance formInstance = form.getFormInstance()
            String formName = eci.resourceFacade.expandNoL10n(formNode.attribute("name"), null)

            WriteSheet sheet = EasyExcel.writerSheet(formName).build()
            FormListEasyExcelRender fler = new FormListEasyExcelRender(formInstance, eci)
            fler.renderSheet(excelWriter, sheet)
        }

        // write file to stream
        excelWriter.finish()
        os.close()
    }

    FormListEasyExcelRender(ScreenForm.FormInstance formInstance, ExecutionContextImpl eci) {
        this.formInstance = formInstance
        this.eci = eci
    }

    void render(OutputStream os) {
        MNode formNode = formInstance.formNode
        String formName = eci.resourceFacade.expandNoL10n(formNode.attribute("name"), null)

        ExcelWriter excelWriter = EasyExcel.write(os).build()
        WriteSheet sheet = EasyExcel.writerSheet(formName).build()
        // render to the sheet
        renderSheet(excelWriter, sheet)
        // write file to stream
        excelWriter.finish()
        os.close()
    }

    enum StyleKey {
        Default, DateTime, Date, Time, Number, Right, YearMonth, NumberRight, Currency

    }

    @CompileStatic
    class StyleStrategy extends AbstractCellStyleStrategy implements RowWriteHandler {

        EnumMap<StyleKey, CellStyle> styleMap
        CellStyle headerStyle
        private int maxRows

        StyleStrategy(int maxRows) {
            this.maxRows = maxRows
        }

        @Override
        protected void initCellStyle(Workbook workbook) {
            this.headerStyle = makeHeaderStyle(workbook)
            CellStyle rowDefaultStyle = makeRowDefaultStyle(workbook)
            CellStyle rowDateTimeStyle = makeRowDateTimeStyle(workbook)
            CellStyle rowDateStyle = makeRowDateStyle(workbook)
            CellStyle rowTimeStyle = makeRowTimeStyle(workbook)
            CellStyle rowNumberStyle = makeRowNumberStyle(workbook)
            CellStyle rowRightStyle = makeRowRightStyle(workbook)
            CellStyle rowYearMonthStyle = makeRowYearMonthStyle(workbook)
            CellStyle rowNumberRightStyle = makeRowNumberRightStyle(workbook)
            CellStyle rowCurrencyStyle = makeRowCurrencyStyle(workbook)
            this.styleMap = new EnumMap(StyleKey);
            styleMap.put(StyleKey.Default, rowDefaultStyle)
            styleMap.put(StyleKey.DateTime, rowDateTimeStyle)
            styleMap.put(StyleKey.Date, rowDateStyle)
            styleMap.put(StyleKey.Time, rowTimeStyle)
            styleMap.put(StyleKey.Number, rowNumberStyle)
            styleMap.put(StyleKey.NumberRight, rowNumberRightStyle)
            styleMap.put(StyleKey.Currency, rowCurrencyStyle)
            styleMap.put(StyleKey.Right, rowRightStyle)
            styleMap.put(StyleKey.YearMonth, rowYearMonthStyle)
        }

        @Override
        void afterCellDispose(WriteSheetHolder writeSheetHolder, WriteTableHolder writeTableHolder, List<CellData> cellDataList, Cell cell, Head head, Integer relativeRowIndex, Boolean isHead) {
            if (isHead) {
                cell.setCellStyle(headerStyle)
            } else {
                if (cellDataList.isEmpty() == false) {
                    CellData data = cellDataList.first()
                    if (data.type != CellDataTypeEnum.EMPTY) {
                        cell.setCellStyle(styleMap.get(((FormListData)data.getData()).style))
                    }
                }
            }
        }

        @Override
        protected void setHeadCellStyle(Cell cell, Head head, Integer relativeRowIndex) {
            //NOOP
        }

        @Override
        protected void setContentCellStyle(Cell cell, Head head, Integer relativeRowIndex) {
            //NOOP
        }

        @Override
        void beforeRowCreate(WriteSheetHolder writeSheetHolder, WriteTableHolder writeTableHolder, Integer rowIndex, Integer relativeRowIndex, Boolean isHead) {

        }

        @Override
        void afterRowCreate(WriteSheetHolder writeSheetHolder, WriteTableHolder writeTableHolder, Row row, Integer relativeRowIndex, Boolean isHead) {

        }

        @Override
        void afterRowDispose(WriteSheetHolder writeSheetHolder, WriteTableHolder writeTableHolder, Row row, Integer relativeRowIndex, Boolean isHead) {
            if (isHead) {
                writeSheetHolder.cachedSheet.createFreezePane(0, 1)
            }
            row.setHeight((short) -1)
        }

    }

    void renderSheet(ExcelWriter excelWriter, WriteSheet sheet) {
        // get form-list data
        ScreenForm.FormListRenderInfo formListInfo = formInstance.makeFormListRenderInfo()
        // MNode formNode = formListInfo.formNode
        ArrayList<ArrayList<MNode>> formListColumnList = formListInfo.getAllColInfo()
        int numColumns = formListColumnList.size()


//        int rowNum = 0

        // ========== header row
//        XSSFRow headerRow = sheet.createRow(rowNum++)
//        int headerColIndex = 0
        List<String> headerTitleList = []
        for (int i = 0; i < numColumns; i++) {
            ArrayList<MNode> columnFieldList = (ArrayList<MNode>) formListColumnList.get(i)
            for (int j = 0; j < columnFieldList.size(); j++) {
                MNode fieldNode = (MNode) columnFieldList.get(j)
                //XSSFCell headerCell = headerRow.createCell(headerColIndex++)
                //headerCell.setCellStyle(headerStyle)

                MNode headerField = fieldNode.first("header-field")
                MNode defaultField = fieldNode.first("default-field")
                String headerTitle = headerField != null ? headerField.attribute("title") : null
                if (headerTitle == null) headerTitle = defaultField != null ? defaultField.attribute("title") : null
                if (headerTitle == null) headerTitle = StringUtilities.camelCaseToPretty(fieldNode.attribute("name"))
                // headerCell.setCellValue(eci.resourceFacade.expand(headerTitle, null))
                headerTitleList.add(eci.resourceFacade.expand(headerTitle, null))
            }
        }
        // should be ArrayList<Map<String, Object>> from call to AggregationUtil.aggregateList()
        Object listObject = formListInfo.getListObject(false)
        ArrayList<Map<String, Object>> listArray
        if (listObject instanceof ArrayList) {
            listArray = (ArrayList<Map<String, Object>>) listObject
        } else {
            throw new IllegalArgumentException("List object from FormListRenderInfo was not an ArrayList as expected, is ${listObject?.getClass()?.getName()}")
        }
        int listArraySize = listArray.size()
        List<List<FormListData>> listData = []
        for (int listIdx = 0; listIdx < listArraySize; listIdx++) {
            Map<String, Object> curRowMap = (Map<String, Object>) listArray.get(listIdx)

            eci.contextStack.push(curRowMap)
            try {
                List<FormListData> rowData = []
                listData.add(rowData)

                for (int i = 0; i < numColumns; i++) {
                    ArrayList<MNode> columnFieldList = (ArrayList<MNode>) formListColumnList.get(i)
                    for (int j = 0; j < columnFieldList.size(); j++) {
                        MNode fieldNode = (MNode) columnFieldList.get(j)
                        // String fieldName = fieldNode.attribute("name")
                        String fieldAlign = fieldNode.attribute("align")

                        MNode defaultField = fieldNode.first("default-field")
                        ArrayList<MNode> childList = defaultField.getChildren()
                        int childListSize = childList.size()

                        if (childListSize == 1) {
                            // use data specific cell type and style
                            MNode widgetNode = (MNode) childList.get(0)
                            String widgetType = widgetNode.getName()
                            String widgetFormat = widgetNode.attribute("format")

                            Object curValue = getFieldValue(fieldNode, widgetNode)

                            // cell type options are _NONE, BLANK, BOOLEAN, ERROR, FORMULA, NUMERIC, STRING
                            // cell value options are: boolean, Date, Calendar, double, String, RichTextString
                            if (curValue instanceof String) {
                                rowData.add(new FormListData(curValue, CellType.STRING, StyleKey.Default))
                            } else if (curValue instanceof Number) {
                                def data = new FormListData(((Number) curValue).doubleValue(), CellType.NUMERIC, StyleKey.Default)
                                rowData.add(data)
                                String currencyUnitField = widgetNode.attribute("currency-unit-field")
                                if (currencyUnitField != null && !currencyUnitField.isEmpty()) {
                                    data.setStyle(StyleKey.Currency)
                                } else if (widgetFormat != null && widgetFormat.contains(".00")) {
                                    data.setStyle(StyleKey.Currency)
                                } else {
                                    if ("right".equals(fieldAlign)) {
                                        if (widgetFormat !=null) {
                                            data.setStyle(StyleKey.NumberRight)
                                        } else {
                                            data.setStyle(StyleKey.Right)
                                        }
                                    }
                                }
                            } else if (curValue instanceof Date) {
                                def data = new FormListData(DateUtil.getExcelDate(curValue), CellType.NUMERIC, StyleKey.Default)
                                rowData.add(data)
                                if (widgetFormat != null && (widgetFormat in ['yyyy/MM', 'yyyy-MM', 'yyyy.MM'])) {
                                    data.setStyle(StyleKey.YearMonth)
                                } else if (curValue instanceof java.sql.Date) {
                                    data.setStyle(StyleKey.Date)
                                } else if (curValue instanceof java.sql.Time) {
                                    data.setStyle(StyleKey.Time)
                                } else {
                                    data.setStyle(StyleKey.DateTime)
                                }
                            } else if (curValue != null) {
                                rowData.add(new FormListData(curValue.toString(), CellType.STRING, StyleKey.Default))
                            } else {
                                rowData.add(null)
                            }

                        } else {
                            // always use string with values from all child elements
                            StringBuilder cellSb = new StringBuilder()

                            for (int childIdx = 0; childIdx < childListSize; childIdx++) {
                                MNode widgetNode = (MNode) childList.get(childIdx)
                                Object curValue = getFieldValue(fieldNode, widgetNode)
                                if (curValue != null) {
                                    if (curValue instanceof String) {
                                        cellSb.append((String) curValue)
                                    } else {
                                        String format = widgetNode.attribute("format")
                                        cellSb.append(eci.l10nFacade.format(curValue, format))
                                    }
                                }
                                if (childIdx < (childListSize - 1) && cellSb.length() > 0 && cellSb.charAt(cellSb.length() - 1) != (char) '\n')
                                    cellSb.append('\n')
                            }

                            rowData.add(new FormListData(cellSb.toString(), CellType.STRING, StyleKey.Default))
                        }

                    }
                }
            } finally {
                eci.contextStack.pop()
            }
        }

        WriteTable writeTable = EasyExcel.writerTable(0).head(headerTitleList.collect({ [it] }))
                .registerWriteHandler( new StyleStrategy(listData.size()) )
                .registerWriteHandler(new LongestMatchColumnWidthStyleStrategy()).registerConverter(new Converter() {
            @Override
            Class supportJavaTypeKey() {
                return FormListData.class
            }

            @Override
            CellDataTypeEnum supportExcelTypeKey() {
                return null
            }

            @Override
            Object convertToJavaData(CellData cellData, ExcelContentProperty contentProperty, GlobalConfiguration globalConfiguration) throws Exception {
                return null
            }

            @Override
            CellData convertToExcelData(Object value, ExcelContentProperty contentProperty, GlobalConfiguration globalConfiguration) throws Exception {
                FormListData origData = value as FormListData
                CellData cellData
                if (origData == null) {
                    cellData = CellData.newEmptyInstance()
                } else if(origData.type == CellType.STRING) {
                    cellData = new CellData((String) origData.value)
                } else if (origData.type == CellType.NUMERIC) {
                    cellData = new CellData((BigDecimal) origData.value)
                } else {
                    throw new RuntimeException("unsupported cell type")
                }
                cellData.setData(origData)
                return cellData
            }
        }).build()

        excelWriter.write(listData, sheet, writeTable)

        // TODO? something special for footer row

        // auto size columns
//        for (int c = 0; c < sheetColCount; c++) sheet.autoSizeColumn(c)
    }

    Object getFieldValue(MNode fieldNode, MNode widgetNode) {
        String widgetType = widgetNode.getName()
        if (ignoreFieldElements.contains(widgetType)) return null
        String fieldName = fieldNode.attribute("name")

        // similar logic to widget types in DefaultScreenMacros.csv.ftl
        // view oriented: link, display, display-entity
        // edit oriented: check, drop-down, radio, date-time, text-area, text-line, text-find

        String conditionAttr = widgetNode.attribute("condition")
        if (conditionAttr != null && !conditionAttr.isEmpty() && !eci.resourceFacade.condition(conditionAttr, null))
            return null

        Object value = null
        if ("display".equals(widgetType) || "display-entity".equals(widgetType) || "link".equals(widgetType) || "label".equals(widgetType)) {
            String entityName = widgetNode.attribute("entity-name")
            String textAttr = widgetNode.attribute("text")

            if (entityName != null && !entityName.isEmpty()) {
                Object fieldValue = eci.contextStack.getByString(fieldName)
                if (fieldValue == null) return getDefaultText(widgetNode)
                EntityDefinition ed = eci.entityFacade.getEntityDefinition(entityName)

                // find the entity value
                String keyFieldName = widgetNode.attribute("key-field-name")
                if (keyFieldName == null || keyFieldName.isEmpty()) keyFieldName = widgetNode.attribute("entity-key-name")
                if (keyFieldName == null || keyFieldName.isEmpty()) keyFieldName = ed.getPkFieldNames().get(0)
                String useCache = widgetNode.attribute("use-cache") ?: widgetNode.attribute("entity-use-cache") ?: "true"

                EntityValue ev = eci.entityFacade.find(entityName).condition(keyFieldName, fieldValue)
                        .useCache(useCache == "true").one()
                if (ev == null) return getDefaultText(widgetNode)

                if (textAttr != null && textAttr.length() > 0) {
                    value = eci.resourceFacade.expand(textAttr, null, ev.getMap())
                } else {
                    // get the value of the default description field for the entity
                    String defaultDescriptionField = ed.getDefaultDescriptionField()
                    if (defaultDescriptionField) value = ev.get(defaultDescriptionField)
                }
            } else if (textAttr != null && !textAttr.isEmpty()) {
                String textMapAttr = widgetNode.attribute("text-map")
                Object textMap = textMapAttr != null ? eci.resourceFacade.expression(textMapAttr, null) : null
                if (textMap instanceof Map) {
                    value = eci.resourceFacade.expand(textAttr, null, textMap)
                } else {
                    value = eci.resourceFacade.expand(textAttr, null)
                }
                if (value == "null") {
                    value = null
                }
            } else {
                value = eci.contextStack.getByString(fieldName)
            }
            if (value == null) {
                String defaultText = widgetNode.attribute("default-text")
                if (defaultText != null && defaultText.length() > 0)
                    value = eci.resourceFacade.expand(defaultText, null)
            }

            /* FUTURE: widget types for interactive form-list, low priority for intended use
        } else if ("drop-down".equals(widgetType) || "check".equals(widgetType) || "radio".equals(widgetType)) {
        } else if ("text-area".equals(widgetType) || "text-line".equals(widgetType) || "text-find".equals(widgetType)) {
         */
        } else {
            value = eci.contextStack.getByString(fieldName)
        }

        return value
    }

    protected String getDefaultText(MNode widgetNode) {
        String defaultText = widgetNode.attribute("default-text")
        if (defaultText != null && defaultText.length() > 0) {
            return eci.resourceFacade.expand(defaultText, null)
        } else {
            return null
        }
    }

    CellStyle makeHeaderStyle(Workbook wb) {
        Font headerFont = wb.createFont()
        headerFont.setBold(true)

        CellStyle style = createNoBorderStyle(wb)
        style.setAlignment(HorizontalAlignment.CENTER)
        style.setFillForegroundColor(IndexedColors.GREY_40_PERCENT.getIndex())
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND)
        style.setFont(headerFont)

        return style
    }

    CellStyle makeRowDefaultStyle(Workbook wb) {
        Font rowFont = wb.createFont()

        CellStyle style = createNoBorderStyle(wb)
        style.setAlignment(HorizontalAlignment.LEFT)
        style.setFont(rowFont)
        style.setWrapText(true)

        return style
    }

    CellStyle makeRowYearMonthStyle(Workbook wb) {
        DataFormat df = wb.createDataFormat()
        Font rowFont = wb.createFont()

        CellStyle style = createNoBorderStyle(wb)
        style.setAlignment(HorizontalAlignment.LEFT)
        style.setFont(rowFont)
        style.setDataFormat(df.getFormat(DateFormatConverter.convert(Locale.US, "yyyy/MM")))

        return style
    }
    CellStyle makeRowDateTimeStyle(Workbook wb) {
        DataFormat df = wb.createDataFormat()
        Font rowFont = wb.createFont()

        CellStyle style = createNoBorderStyle(wb)
        style.setAlignment(HorizontalAlignment.LEFT)
        style.setFont(rowFont)
        style.setDataFormat(df.getFormat(DateFormatConverter.convert(Locale.US, "yyyy-MM-dd HH:mm")))

        return style
    }

    CellStyle makeRowDateStyle(Workbook wb) {
        DataFormat df = wb.createDataFormat()
        Font rowFont = wb.createFont()

        CellStyle style = createNoBorderStyle(wb)
        style.setAlignment(HorizontalAlignment.LEFT)
        style.setFont(rowFont)
        style.setDataFormat(df.getFormat(DateFormatConverter.convert(Locale.US, "yyyy-MM-dd")))

        return style
    }

    CellStyle makeRowTimeStyle(Workbook wb) {
        DataFormat df = wb.createDataFormat()
        Font rowFont = wb.createFont()

        CellStyle style = createNoBorderStyle(wb)
        style.setAlignment(HorizontalAlignment.LEFT)
        style.setFont(rowFont)
        style.setDataFormat(df.getFormat(DateFormatConverter.convert(Locale.US, "HH:mm:ss")))

        return style
    }

    CellStyle makeRowNumberStyle(Workbook wb) {
        DataFormat df = wb.createDataFormat()
        Font rowFont = wb.createFont()

        CellStyle style = createNoBorderStyle(wb)
        style.setAlignment(HorizontalAlignment.CENTER)
        style.setFont(rowFont)
        style.setDataFormat(df.getFormat("#,##0.0######"))

        return style
    }

    CellStyle makeRowRightStyle(Workbook wb) {
        Font rowFont = wb.createFont()

        CellStyle style = createNoBorderStyle(wb)
        style.setAlignment(HorizontalAlignment.RIGHT)
        style.setFont(rowFont)

        return style
    }
    CellStyle makeRowNumberRightStyle(Workbook wb) {
        DataFormat df = wb.createDataFormat()
        Font rowFont = wb.createFont()

        CellStyle style = createNoBorderStyle(wb)
        style.setAlignment(HorizontalAlignment.RIGHT)
        style.setFont(rowFont)
        style.setDataFormat(df.getFormat("#,##0.0#####"))

        return style
    }

    CellStyle makeRowCurrencyStyle(Workbook wb) {
        DataFormat df = wb.createDataFormat()
        Font rowFont = wb.createFont()

        CellStyle style = createNoBorderStyle(wb)
        style.setAlignment(HorizontalAlignment.RIGHT)
        style.setFont(rowFont)
        style.setDataFormat(df.getFormat('#,##0.00_);[Red](#,##0.00)'))

        return style
    }

    CellStyle createBorderedStyle(Workbook wb) {
        BorderStyle thin = BorderStyle.THIN
        short black = IndexedColors.BLACK.getIndex()

        CellStyle style = wb.createCellStyle()
        style.setBorderRight(thin)
        style.setRightBorderColor(black)
        style.setBorderBottom(thin)
        style.setBottomBorderColor(black)
        style.setBorderLeft(thin)
        style.setLeftBorderColor(black)
        style.setBorderTop(thin)
        style.setTopBorderColor(black)
        return style
    }

    CellStyle createNoBorderStyle(Workbook wb) {
        BorderStyle noBorder = BorderStyle.NONE
        CellStyle style = wb.createCellStyle()
        style.setBorderRight(noBorder)
        style.setBorderBottom(noBorder)
        style.setBorderLeft(noBorder)
        style.setBorderTop(noBorder)
        style.setVerticalAlignment(VerticalAlignment.TOP)
        // logger.warn("created no border style ${style.properties}")
        return style
    }
}
