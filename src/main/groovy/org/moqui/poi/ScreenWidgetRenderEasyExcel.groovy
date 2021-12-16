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
import com.alibaba.excel.write.metadata.WriteSheet
import groovy.transform.CompileStatic
import org.apache.logging.log4j.Level
import org.moqui.impl.context.ExecutionContextImpl
import org.moqui.impl.screen.*
import org.moqui.util.ContextStack
import org.moqui.util.MNode
import org.slf4j.Logger
import org.slf4j.LoggerFactory

/**
 * ScreenWidgetRender implementation to generate Excel (xlsx) file with a sheet for each form-list found on the screen.
 *
 * This can be used with the form-list.@show-xlsx-button attribute.
 *
 * To specify a link manually in a XML Screen do something like:
 * <pre>
 *         <link url="${sri.getScreenUrlInstance().getScreenOnlyPath()}" url-type="plain" text="XLS Export"
 *                 parameter-map="ec.web.requestParameters + [saveFilename:('FinancialInfo_' + partyId + '.xlsx'),
 *                     renderMode:'xlsx', pageNoLimit:'true', lastStandalone:'true']"/>
 * </pre>
 */
@CompileStatic
class ScreenWidgetRenderEasyExcel implements ScreenWidgetRender {
    private static final Logger logger = LoggerFactory.getLogger(ScreenWidgetRenderEasyExcel.class)
    private static final String excelWriterFieldName = "WidgetRenderEasyExcelWriter"
    static {
        org.apache.logging.log4j.core.config.Configurator.setLevel("org.apache.poi.util.XMLHelper", Level.ERROR);
    }
    ScreenWidgetRenderEasyExcel() { }

    @Override
    void render(ScreenWidgets widgets, ScreenRenderImpl sri) {
        ContextStack cs = sri.ec.contextStack
        cs.push()
        try {
            cs.sri = sri
            OutputStream os = sri.getOutputStream()
            ExcelWriter  excelWriter = (ExcelWriter) cs.getByString(excelWriterFieldName)
            boolean createdWorkbook = false
            if (excelWriter == null) {
                excelWriter = EasyExcel.write(os).build()
                createdWorkbook = true
                cs.put(excelWriterFieldName, excelWriter)
            }

            MNode widgetsNode = widgets.widgetsNode
            if (widgetsNode.name == "screen") widgetsNode = widgetsNode.first("widgets")
            renderSubNodes(widgetsNode, sri, excelWriter)

            if (createdWorkbook) {
                excelWriter.finish()
                os.close()
            }
        } finally {
            cs.pop()
        }
    }

    static void renderSubNodes(MNode widgetsNode, ScreenRenderImpl sri, ExcelWriter excelWriter) {
        ExecutionContextImpl eci = sri.ec
        ScreenDefinition sd = sri.getActiveScreenDef()

        // iterate over child elements to find and render form-list
        // recursive renderSubNodes() call for: container-box (box-body, box-body-nopad), container-row (row-col)
        ArrayList<MNode> childList = widgetsNode.getChildren()
        int childListSize = childList.size()
        for (int i = 0; i < childListSize; i++) {
            MNode childNode = (MNode) childList.get(i)
            String nodeName = childNode.getName()
            if ("form-list".equals(nodeName)) {
                ScreenForm form = sd.getForm(childNode.attribute("name"))
                MNode formNode = form.getOrCreateFormNode()
                String formName = eci.resourceFacade.expandNoL10n(formNode.attribute("name"), null)

                ScreenForm.FormInstance formInstance = form.getFormInstance()

                WriteSheet sheet = EasyExcel.writerSheet(formName).build()
                FormListEasyExcelRender fler = new FormListEasyExcelRender(formInstance, eci)
                fler.renderSheet(excelWriter, sheet)
            } else if ("section".equals(nodeName)) {
                // nest into section by calling renderSection() so conditions, actions are run (skip section-iterate)
                sri.renderSection(childNode.attribute("name"))
            } else if ("container-box".equals(nodeName)) {
                MNode boxBody = childNode.first("box-body")
                if (boxBody != null) renderSubNodes(boxBody, sri, excelWriter)
                MNode boxBodyNopad = childNode.first("box-body-nopad")
                if (boxBodyNopad != null) renderSubNodes(boxBodyNopad, sri, excelWriter)
            } else if ("container-row".equals(nodeName)) {
                MNode rowCol = childNode.first("row-col")
                if (rowCol != null) renderSubNodes(rowCol, sri, excelWriter)
            } else if ("container".equals(nodeName)) {
                renderSubNodes(childNode, sri, excelWriter)
            }
            // NOTE: other elements ignored, including section-iterate (outside intended use case for Excel render
        }
    }
}
