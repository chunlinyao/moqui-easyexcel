package org.moqui.poi

import groovy.transform.Canonical
import groovy.transform.CompileStatic
import org.apache.poi.ss.usermodel.CellType

@Canonical
@CompileStatic
class FormListData {
    Object value
    CellType type
    FormListEasyExcelRender.StyleKey style
}
