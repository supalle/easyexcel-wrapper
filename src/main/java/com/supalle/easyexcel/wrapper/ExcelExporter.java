package com.supalle.easyexcel.wrapper;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.ExcelWriter;
import com.alibaba.excel.write.builder.ExcelWriterSheetBuilder;
import com.alibaba.excel.write.handler.AbstractRowWriteHandler;
import com.alibaba.excel.write.metadata.WriteSheet;
import com.alibaba.excel.write.metadata.holder.WriteSheetHolder;
import com.alibaba.excel.write.metadata.holder.WriteTableHolder;
import com.alibaba.excel.write.style.HorizontalCellStyleStrategy;
import lombok.Getter;
import org.apache.poi.ss.usermodel.Comment;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.springframework.util.CollectionUtils;

import java.io.File;
import java.io.OutputStream;
import java.util.*;
import java.util.function.Function;
import java.util.stream.Collectors;

import static com.supalle.easyexcel.wrapper.Common.isNotEmpty;

@Getter
public class ExcelExporter {

    private final List<Sheet> sheets = new ArrayList<>();

    private File outFile;
    private OutputStream outputStream;
    private Function<String, Dict> defaultDictSupplier;
    private boolean autoCloseStream = true;

    public static ExcelExporter create() {
        return new ExcelExporter();
    }

    public ExcelExporter outFile(File outFile) {
        this.outFile = outFile;
        return this;
    }

    public ExcelExporter outputStream(OutputStream outputStream) {
        this.outputStream = outputStream;
        return this;
    }

    public ExcelExporter defaultDictSupplier(Function<String, Dict> defaultDictSupplier) {
        this.defaultDictSupplier = defaultDictSupplier;
        return this;
    }

    public ExcelExporter autoCloseStream(boolean autoCloseStream) {
        this.autoCloseStream = autoCloseStream;
        return this;
    }

    public <T> Sheet<T> sheet(String sheetName) {
        return sheet(sheetName, null);
    }

    public <T> Sheet<T> sheet(ExcelEntity<T> excelEntity) {
        return sheet(null, excelEntity);
    }

    public <T> Sheet<T> sheet(String sheetName, ExcelEntity<T> excelEntity) {
        Sheet<T> sheet = new Sheet<>();
        this.sheets.add(sheet);
        return sheet.parent(this).excelEntity(excelEntity).sheetName(sheetName == null ? nextSheetName() : sheetName);
    }

    public void letItGo() {
        startImport();
    }

    public void startImport() {

        if (this.sheets.isEmpty()) {
            throw new ExcelException("??????????????????????????????(Sheet)???");
        }
        for (int i = 0; i < this.sheets.size(); i++) {
            Sheet sheet = this.sheets.get(i);
            if (sheet.getExcelEntity() == null) {
                throw new ExcelException(String.format("???%d??????????????????????????????Null", i));
            }
            List<ExcelEntity.ExcelColumnMapping> excelColumnMappings = sheet.getExcelEntity().getExcelColumnMappings();
            if (excelColumnMappings != null && !excelColumnMappings.isEmpty()) {
                for (ExcelEntity.ExcelColumnMapping excelColumnMapping : excelColumnMappings) {
                    if (excelColumnMapping.getExcelColumnExportMapping() == null) {
                        throw new ExcelException(String.format("???%d???????????????????????????%s????????????getting??????", i, excelColumnMapping.getHeadName()));
                    }
                }
            }
        }

        ExcelWriter excelWriter = null;
        try {

            if (outputStream != null) {
                if (autoCloseStream) {
                    excelWriter = EasyExcel.write(outputStream).autoCloseStream(true).build();
                } else {
                    excelWriter = EasyExcel.write(outputStream).autoCloseStream(false).build();
                }

            } else {
                if (outFile == null) {
                    throw new ExcelException("?????????????????????Excel????????????????????????");
                }

                if (!outFile.exists()) {
                    throw new ExcelException("???????????????Excel??????????????? " + outFile.getName());
                }

                if (outFile.isDirectory()) {
                    throw new ExcelException("??????????????????????????? " + outFile.getName());
                }
                excelWriter = EasyExcel.write(outFile).autoCloseStream(true).build();
            }
            doWrite(excelWriter);
        } finally {
            // ???????????????finish ??????????????????
            if (excelWriter != null) {
                excelWriter.finish();
            }
        }
    }

    // TODO ?????????-?????????
    private void doWrite(ExcelWriter excelWriter) {
        int sheetNo = 0;
        for (Sheet sheet : this.sheets) {

            Map<String, Dict> dictMap = new HashMap<>();
            Map<String, Map<String, Dict.DictItem>> dictItemMap = new HashMap<>();

            ExcelEntity excelEntity = sheet.getExcelEntity();
            HorizontalCellStyleStrategy horizontalCellStyleStrategy = excelEntity.getHorizontalCellStyleStrategy();
            List<ExcelEntity.ExcelColumnMapping> excelColumnMappings = excelEntity.getExcelColumnMappings();

            List<List<String>> headList = new ArrayList<>();
            for (ExcelEntity.ExcelColumnMapping mapping : excelColumnMappings) {
                // ???????????????
                String headName = String.valueOf(mapping.getHeadName()).trim();
                headList.add(Collections.singletonList(headName));

                // ????????????
                String dictName = mapping.getDict();
                if (dictName != null) {
                    Function<String, Dict> dictSupplier = sheet.getDictSupplier() == null ? sheet.getParent().getDefaultDictSupplier() : sheet.getDictSupplier();
                    if (Optional.ofNullable(dictSupplier).isPresent()) {
                        Dict dict = dictSupplier.apply(dictName);
                        if (dict == null) {
                            throw new ExcelException(String.format("??????'%s'????????????", dictName));
                        }
                        dictMap.put(dictName, dict);
                        Map<String, Dict.DictItem> tempDictItemMap = Optional.ofNullable(dict.getDictItems())
                                .orElseGet(ArrayList::new)
                                .stream().collect(Collectors.toMap(Dict.DictItem::getValue, e -> e));
                        dictItemMap.put(dictName, tempDictItemMap);
                    } else {
                        throw new ExcelException(String.format("??????????????????'%s'???????????????", dictName));
                    }
                }
            }

            List<List<Object>> data = new LinkedList<>();

            if (!CollectionUtils.isEmpty(sheet.getData())) {
                for (Object datum : sheet.getData()) {
                    List<Object> cellValues = new ArrayList<>();
                    data.add(cellValues);
                    if (datum == null) {
                        for (int i = 0; i < excelColumnMappings.size(); i++) {
                            cellValues.add(null);
                        }
                        continue;
                    }

                    for (ExcelEntity.ExcelColumnMapping mapping : excelColumnMappings) {
                        ExcelEntity.ExcelColumnExportMapping excelColumnExportMapping = mapping.getExcelColumnExportMapping();

                        Function getting = excelColumnExportMapping.getGetting();
                        Object cellValue = null;
                        if (getting != null) {
                            cellValue = getting.apply(datum);
                        }
                        Function formatter = excelColumnExportMapping.getFormatter();
                        if (formatter != null) {
                            cellValue = formatter.apply(cellValue);
                        }

                        if (excelColumnExportMapping.isJumpNull() && cellValue == null) {
                            cellValues.add(null);
                            continue;
                        }
                        if (excelColumnExportMapping.isJumpEmpty() && (cellValue == null || cellValue.toString().trim().isEmpty())) {
                            cellValues.add(null);
                            continue;
                        }

                        String dictName = mapping.getDict();
                        if (dictName != null) {
                            Map<String, Dict.DictItem> itemMap = dictItemMap.get(dictName);
                            Dict.DictItem dictItem = null;
                            if ((dictItem = itemMap.get(String.valueOf(cellValue))) == null) {
                                Dict dict = dictMap.get(dictName);
                                throw new ExcelException(String.format("???'%s'????????????'%s'?????????????????????????????????'%s:%s'?????????%s", mapping.getHeadName(), cellValue, dictName, dict.getComment(), dict.getDictItems().toString()));
                            }
                            if (!mapping.isDictUsedValue()) {
                                cellValue = dictItem.getLabel();
                            }
                        }

                        if (excelColumnExportMapping.isAutoTrim()) {
                            if (cellValue instanceof String) {
                                cellValue = cellValue.toString().trim();
                            }
                        }
                        cellValues.add(cellValue);
                    }
                }
            }

            ExcelWriterSheetBuilder excelWriterSheetBuilder = EasyExcel.writerSheet(sheetNo, sheet.getSheetName() == null ? "sheet" + sheetNo : sheet.getSheetName())
                    .head(headList);
            if (horizontalCellStyleStrategy != null) {
                excelWriterSheetBuilder = excelWriterSheetBuilder.registerWriteHandler(horizontalCellStyleStrategy);
            }
            excelWriterSheetBuilder = excelWriterSheetBuilder.registerWriteHandler(new AbstractRowWriteHandler() {

                public void afterRowDispose(WriteSheetHolder writeSheetHolder, WriteTableHolder writeTableHolder, Row row,
                                            Integer relativeRowIndex, Boolean isHead) {
                    if (isHead) {
                        org.apache.poi.ss.usermodel.Sheet sheet = writeSheetHolder.getSheet();
                        Drawing<?> drawingPatriarch = sheet.createDrawingPatriarch();
                        int cellIndex = 0;
                        for (ExcelEntity.ExcelColumnMapping mapping : excelColumnMappings) {
                            String columnComment = mapping.getComment();
                            String dict = mapping.getDict();
                            if (isNotEmpty(columnComment) || isNotEmpty(dict)) {

                                List<String> comments = new ArrayList<>();
                                if (isNotEmpty(columnComment)) {
                                    comments.add("?????????" + columnComment);
                                }
                                if (isNotEmpty(dict)) {
                                    Dict d = dictMap.get(dict);
                                    comments.add(String.format("?????????%s[%s]", d.getComment(), mapping.isDictUsedValue() ? "???" : "??????"));
                                    List<Dict.DictItem> dictItems = d.getDictItems();
                                    comments.add("    ???:??????    ");
                                    for (Dict.DictItem dictItem : dictItems) {
                                        comments.add(String.format("    [%s]:[%s]%s    ", dictItem.getValue(), dictItem.getLabel(), dictItem.getComment() == null ? "" : dictItem.getComment()));
                                    }
                                }
                                // ???????????? ???????????????????????????
                                Comment comment = drawingPatriarch
                                        .createCellComment(new XSSFClientAnchor(0, 0, 0, 0, (short) 1, 0, (short) 2, 1));

                                // ??????????????????
                                comment.setString(new XSSFRichTextString(String.join("\n", comments)));
                                // ????????????????????????????????????
                                sheet.getRow(0).getCell(cellIndex).setCellComment(comment);
                            }
                            cellIndex++;
                        }
                    }
                }

            });
            WriteSheet writeSheet = excelWriterSheetBuilder.build();
            excelWriter.write(data, writeSheet);

        }

    }


    private String nextSheetName() {
        return "sheet" + this.sheets.size();
    }

    @Getter
    public static class Sheet<T> {
        private ExcelExporter parent;

        private String sheetName;
        private ExcelEntity<T> excelEntity;
        private Function<String, Dict> dictSupplier;
        private List<T> data;

        public Sheet<T> parent(ExcelExporter parent) {
            this.parent = parent;
            return this;
        }

        public Sheet<T> sheetName(String sheetName) {
            this.sheetName = sheetName;
            return this;
        }

        public <E> Sheet<E> excelEntity(ExcelEntity<E> excelEntity) {
            this.parent.sheets.remove(this);
            Sheet<E> sheet = new Sheet<>();
            this.parent.sheets.add(sheet);
            sheet.parent = this.parent;
            sheet.sheetName = this.sheetName;
            sheet.excelEntity = excelEntity;
            sheet.dictSupplier = this.dictSupplier;
            sheet.data = (List<E>) this.data;
            return sheet;
        }

        public Sheet<T> dictSupplier(Function<String, Dict> dictSupplier) {
            this.dictSupplier = dictSupplier;
            return this;
        }

        public Sheet<T> data(List<T> data) {
            this.data = data;
            return this;
        }

        public Sheet<T> copy() {
            return copyAndRename(this.parent.nextSheetName());
        }

        public Sheet<T> copyAndRename(String sheetName) {
            Sheet<T> last = this;
            Sheet<T> sheet = new Sheet<>();
            this.parent.sheets.add(sheet);
            return sheet.parent(this.parent).sheetName(sheetName).excelEntity(last.getExcelEntity()).dictSupplier(last.getDictSupplier()).data(last.getData());
        }

        public void letItGo() {
            startImport();
        }

        public void startImport() {
            this.parent.startImport();
        }
    }

}
