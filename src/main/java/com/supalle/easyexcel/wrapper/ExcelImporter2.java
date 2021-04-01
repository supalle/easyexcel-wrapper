package com.supalle.easyexcel.wrapper;

import cn.hutool.core.convert.Convert;
import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.context.AnalysisContext;
import com.alibaba.excel.event.AnalysisEventListener;
import com.alibaba.excel.read.listener.ReadListener;
import com.alibaba.excel.read.metadata.ReadSheet;
import lombok.Getter;
import lombok.extern.slf4j.Slf4j;

import java.io.File;
import java.io.InputStream;
import java.lang.reflect.Constructor;
import java.util.*;
import java.util.function.Consumer;
import java.util.function.Function;
import java.util.function.Supplier;
import java.util.regex.Pattern;
import java.util.stream.Collectors;

@Slf4j
@Getter
public class ExcelImporter2 {

    private final List<Sheet> sheets = new ArrayList<>();

    private File inFile;
    private InputStream inputStream;
    private boolean autoCloseStream = true;
    private Function<String, Dict> defaultDictSupplier;

    public static ExcelImporter2 create() {
        return new ExcelImporter2();
    }

    public ExcelImporter2 inFile(File inFile) {
        this.inFile = inFile;
        return this;
    }

    public ExcelImporter2 inputStream(InputStream inputStream) {
        this.inputStream = inputStream;
        return this;
    }

    public ExcelImporter2 defaultDictSupplier(Function<String, Dict> defaultDictSupplier) {
        this.defaultDictSupplier = defaultDictSupplier;
        return this;
    }

    public ExcelImporter2 autoCloseStream(boolean autoCloseStream) {
        this.autoCloseStream = autoCloseStream;
        return this;
    }

    public <T> Sheet<T> sheet() {
        return sheet(null, null);
    }

    public <T> Sheet<T> sheet(String sheetNamePattern) {
        return sheet(sheetNamePattern, null);
    }

    public <T> Sheet<T> sheet(ExcelEntity<T> excelEntity) {
        return sheet(null, excelEntity);
    }

    public <T> Sheet<T> sheet(String sheetNamePattern, ExcelEntity<T> excelEntity) {
        Sheet<T> sheet = new Sheet<>();
        this.sheets.add(sheet);
        return sheet.parent(this).excelEntity(excelEntity).sheetNamePattern(sheetNamePattern == null ? null : Pattern.compile(sheetNamePattern));
    }


    public void letItGo() {
        startExport();
    }

    public void startExport() {

        if (this.sheets.isEmpty()) {
            throw new ExcelException("请添加要读取的工作表(Sheet)。");
        }
        for (int i = 0; i < this.sheets.size(); i++) {
            Sheet sheet = this.sheets.get(i);
            if (sheet.block < 1) {
                sheet.block = 1;
            }
            if (sheet.getExcelEntity() == null) {
                throw new ExcelException(String.format("第%d张工作表的实体映射为Null", i));
            }
            List<ExcelEntity.ExcelColumnMapping> excelColumnMappings = sheet.getExcelEntity().getExcelColumnMappings();
            if (excelColumnMappings != null && !excelColumnMappings.isEmpty()) {
                for (ExcelEntity.ExcelColumnMapping excelColumnMapping : excelColumnMappings) {
                    if (excelColumnMapping.getExcelColumnImportMapping() == null) {
                        throw new ExcelException(String.format("第%d张工作表的映射字段%s缺少导入setting操作", i, excelColumnMapping.getHeadName()));
                    }
                }
            }
        }

        if (inputStream != null) {
            if (autoCloseStream) {
                EasyExcel.read(inputStream).autoCloseStream(true).build().read(buildReadSheets());
            } else {
                EasyExcel.read(inputStream).autoCloseStream(false).build().read(buildReadSheets());
            }
            return;
        }

        if (inFile == null) {
            throw new ExcelException("需要指定导入的Excel文件或者文件流。");
        }

        if (!inFile.exists()) {
            throw new ExcelException("指定导入的Excel文件不存在 " + inFile.getName());
        }

        if (inFile.isDirectory()) {
            throw new ExcelException("无法导入一个文件夹 " + inFile.getName());
        }

        EasyExcel.read(inFile).autoCloseStream(true).build().read(buildReadSheets());

    }

    private List<ReadSheet> buildReadSheets() {
        List<ReadSheet> list = new ArrayList<>();

        int sheetNo = 0;
        for (Sheet sheet : this.sheets) {
            ReadListener readListener = buildListener(sheet);
            list.add(EasyExcel.readSheet(sheetNo++).registerReadListener(readListener).build());
        }

        return list;
    }

    private ReadListener buildListener(Sheet sheet) {


        return new AnalysisEventListener<Map<Integer, String>>() {

            private List<Row<ExcelEntity.ExcelColumnMapping>> excelColumnMappings;
            private Supplier<?> entitySupplier;

            private Map<String, Dict> dictMap;
            private Map<Integer, Map<String, Dict.DictItem>> dictItemMap;

            private List<Row<?>> rows = new LinkedList<>();

            @Override
            public void invokeHeadMap(Map<Integer, String> headMap, AnalysisContext context) {
                List<ExcelEntity.ExcelColumnMapping> excelColumnMappings = sheet.getExcelEntity().getExcelColumnMappings();
                Map<String, Integer> indexMap = headMap.entrySet().stream().collect(Collectors.toMap(e -> String.valueOf(e.getValue()).trim(), e -> e.getKey()));
                this.excelColumnMappings = new ArrayList<>(excelColumnMappings.size());
                this.dictMap = new HashMap<>();
                this.dictItemMap = new HashMap<>();
                for (ExcelEntity.ExcelColumnMapping mapping : excelColumnMappings) {
                    // 查找列索引
                    String headName = String.valueOf(mapping.getHeadName()).trim();
                    Integer columnIndex = indexMap.get(headName);
                    if (columnIndex == null) {
                        throw new ExcelException(String.format("必须包含'%s'列", headName));

                    } else {
                        this.excelColumnMappings.add(new Row<>(columnIndex, mapping));
                    }
                    // 查找字典
                    String dictName = mapping.getDict();
                    if (dictName != null) {
                        Function<String, Dict> dictSupplier = sheet.getDictSupplier() == null ? sheet.getParent().getDefaultDictSupplier() : sheet.getDictSupplier();
                        if (Optional.ofNullable(dictSupplier).isPresent()) {
                            Dict dict = dictSupplier.apply(dictName);
                            if (dict == null) {
                                throw new ExcelException(String.format("字典'%s'没有提供", dictName));
                            }
                            dictMap.put(dictName, dict);
                            Map<String, Dict.DictItem> dictItemMap = Optional.ofNullable(dict.getDictItems())
                                    .orElseGet(ArrayList::new)
                                    .stream().collect(Collectors.toMap(e -> mapping.isDictUsedValue() ? e.getValue() : e.getLabel()
                                            , e -> e));
                            this.dictItemMap.put(columnIndex, dictItemMap);
                        } else {
                            throw new ExcelException(String.format("没有提供字典'%s'的获取途径", dictName));
                        }
                    }
                }
                //
                try {
                    Constructor constructor = sheet.getExcelEntity().getEntityClass().getConstructor();
                    entitySupplier = () -> {
                        try {
                            return constructor.newInstance();
                        } catch (Exception e) {
                            log.error(e.getMessage(), e);
                            throw new ExcelException(String.format("实体类型'%s'创建对象失败。", sheet.getExcelEntity().getEntityClass().getName()));
                        }
                    };
                } catch (Exception e) {
                    log.error(e.getMessage(), e);
                    throw new ExcelException(String.format("实体类型'%s'没有空参构造器。", sheet.getExcelEntity().getEntityClass().getName()));
                }

            }

            @Override
            public void invoke(Map<Integer, String> data, AnalysisContext context) {
                Object obj = entitySupplier.get();
                Row<Object> row = new Row<>(context.readRowHolder().getRowIndex(), obj);
                this.rows.add(row);
                for (Row<ExcelEntity.ExcelColumnMapping> tuple : excelColumnMappings) {
                    Integer columnIndex = tuple.getIndex();
                    ExcelEntity.ExcelColumnMapping mapping = tuple.getData();
                    ExcelEntity.ExcelColumnImportMapping importMapping = mapping.getExcelColumnImportMapping();

                    String cellValue = data.get(columnIndex);
                    if (importMapping.isRequired() && (cellValue == null || cellValue.trim().length() == 0)) {
                        throw new ExcelException(String.format("第%d行的'%s'列不能为空", row.getIndex(), mapping.getHeadName()));
                    }
                    if (importMapping.isJumpNull() && cellValue == null) {
                        continue;
                    }
                    if (importMapping.isJumpEmpty() && (cellValue == null || cellValue.trim().length() == 0)) {
                        continue;
                    }
                    if (importMapping.isAutoTrim() && cellValue != null) {
                        cellValue = cellValue.trim();
                    }

                    String dictName = mapping.getDict();
                    Dict.DictItem dictItem = null;
                    if (mapping.getDict() != null) {
                        Dict dict = dictMap.get(dictName);
                        Map<String, Dict.DictItem> itemMap = this.dictItemMap.get(columnIndex);
                        if ((dictItem = itemMap.get(cellValue)) == null) {
                            throw new ExcelException(String.format("列'%s'的字典值'%s'超出约定范围，可用字典'%s:%s'只包含%s", mapping.getHeadName(), cellValue, dictName, dict.getComment(), dict.getDictItems().toString()));
                        }
                        cellValue = dictItem.getValue();
                    }
                    if ((importMapping.getSetting() != null)) {
                        Function<String, ?> formatter = importMapping.getFormatter();
                        if (formatter != null) {
                            if (importMapping.getSetting() != null)
                                importMapping.getSetting().accept(obj, formatter.apply(cellValue));
                        } else {
                            Class<?> type = importMapping.getType();
                            if (Dict.DictItem.class.isAssignableFrom(type)) {
                                importMapping.getSetting().accept(obj, dictItem);
                            } else {
                                importMapping.getSetting().accept(obj, Convert.convert(type, cellValue));
                            }
                        }
                    }
                }
                if (rows.size() % sheet.block == 0) {
                    sheet.getHandler().accept(rows);
                    rows = new LinkedList<>();
                }
            }

            @Override
            public void doAfterAllAnalysed(AnalysisContext context) {
                if (!rows.isEmpty()) {
                    sheet.getHandler().accept(rows);
                    rows = new LinkedList<>();
                }
            }
        };
    }

    private String nextSheetName() {
        return "sheet" + this.sheets.size();
    }

    @Getter
    public static class Sheet<T> {
        private ExcelImporter2 parent;

        private Pattern sheetNamePattern; // TODO
        private ExcelEntity<T> excelEntity;
        private Function<String, Dict> dictSupplier;
        private int block = 1;
        private Consumer<List<Row<T>>> handler;

        public Sheet<T> parent(ExcelImporter2 parent) {
            this.parent = parent;
            return this;
        }

        public Sheet<T> sheetNamePattern(Pattern sheetNamePattern) {
            this.sheetNamePattern = sheetNamePattern;
            return this;
        }

        public Sheet<T> block(int block) {
            this.block = block;
            return this;
        }

        public <E> Sheet<E> excelEntity(ExcelEntity<E> excelEntity) {
            this.parent.sheets.remove(this);
            Sheet<E> sheet = new Sheet<>();
            this.parent.sheets.add(sheet);
            sheet.parent = this.parent;
            sheet.sheetNamePattern = this.sheetNamePattern;
            sheet.excelEntity = excelEntity;
            sheet.dictSupplier = this.dictSupplier;
            return sheet;
        }

        public Sheet<T> handler(Consumer<List<Row<T>>> handler) {
            this.handler = handler;
            return this;
        }

        public Sheet<T> dictSupplier(Function<String, Dict> dictSupplier) {
            this.dictSupplier = dictSupplier;
            return this;
        }

        public Sheet<T> copy() {
            Sheet<T> last = this;
            Sheet<T> sheet = new Sheet<>();
            this.parent.sheets.add(sheet);
            return sheet.parent(this.parent).sheetNamePattern(last.sheetNamePattern)
                    .excelEntity(last.getExcelEntity()).dictSupplier(last.getDictSupplier()).block(last.getBlock()).handler(last.getHandler());
        }

        public void letItGo() {
            startExport();
        }

        public void startExport() {
            this.parent.startExport();
        }
    }

}
