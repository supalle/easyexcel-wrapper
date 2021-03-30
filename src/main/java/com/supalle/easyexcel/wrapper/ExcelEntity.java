package com.supalle.easyexcel.wrapper;

import com.alibaba.excel.write.style.HorizontalCellStyleStrategy;
import lombok.Getter;
import lombok.Setter;

import java.util.*;
import java.util.function.BiConsumer;
import java.util.function.Function;

@Getter
public class ExcelEntity<E> {

    private Class<E> entityClass;
    private List<ExcelColumnMapping> excelColumnMappings;
    private HorizontalCellStyleStrategy horizontalCellStyleStrategy;

    public static <E> ExcelEntity<E> of(Class<E> entityClass) {
        ExcelEntity<E> excelEntity = new ExcelEntity<>();
        excelEntity.entityClass = entityClass;
        excelEntity.excelColumnMappings = new ArrayList<>();
        return excelEntity;
    }

    public ExcelEntity<E> horizontalCellStyleStrategy(HorizontalCellStyleStrategy horizontalCellStyleStrategy) {
        this.horizontalCellStyleStrategy = horizontalCellStyleStrategy;
        return this;
    }

    public ExcelColumnMapping<E, String, String> mapping(String headName) {
        return mapping(headName, String.class, String.class);
    }

    public <T> ExcelColumnMapping<E, T, T> mapping(String headName, Class<T> type) {
        return mapping(headName, type, type);
    }

    public <Ex, Im> ExcelColumnMapping<E, Ex, Im> mapping(String headName, Class<Ex> exportType, Class<Im> importType) {
        ExcelColumnMapping<E, Ex, Im> mapping = new ExcelColumnMapping<>();
        mapping.parent = this;
        mapping.headName = headName;
        mapping.excelColumnExportMapping = new ExcelColumnExportMapping<>();
        mapping.excelColumnExportMapping.parent = mapping;
        mapping.excelColumnExportMapping.type(exportType);
        mapping.excelColumnImportMapping = new ExcelColumnImportMapping<>();
        mapping.excelColumnImportMapping.parent = mapping;
        mapping.excelColumnImportMapping.type(importType);
        excelColumnMappings.add(mapping);
        return mapping;
    }

    public List<ExcelColumnOption> toExcelColumnOptions() {
        List<ExcelColumnOption> list = new ArrayList<>();

        if (this.excelColumnMappings == null || this.excelColumnMappings.isEmpty()) {
            return list;
        }

        for (ExcelColumnMapping mapping : this.excelColumnMappings) {
            list.add(ExcelColumnOption.builder().headName(mapping.getHeadName()).required(mapping.getExcelColumnImportMapping() != null && mapping.getExcelColumnImportMapping().required).build());
        }

        return list;
    }

    public ExcelEntity<E> cut(List<ExcelColumnOption> excelColumnOptions) {
        return cut(excelColumnOptions, true);
    }

    public ExcelEntity<E> cut(List<ExcelColumnOption> excelColumnOptions, boolean distinctHeadName) {
        if (excelColumnOptions == null || excelColumnOptions.isEmpty()) {
            return this;
        }
        if (this.excelColumnMappings == null || this.excelColumnMappings.isEmpty()) {
            return this;
        }
        // 去重，优先选择required的
        if (distinctHeadName) {
            LinkedHashMap<String, ExcelColumnOption> newOptions = new LinkedHashMap<>();
            for (ExcelColumnOption option : excelColumnOptions) {
                if (option == null) {
                    continue;
                }
                ExcelColumnOption old = newOptions.get(option.getHeadName());
                if (old != null && Boolean.TRUE.equals(old.getRequired())) {
                    continue;
                }
                newOptions.put(option.getHeadName(), option);
            }
            excelColumnOptions = new ArrayList<>(newOptions.values());
        }

        // 裁剪
        List<ExcelColumnMapping> excelColumnMappings = this.excelColumnMappings;
        List<ExcelColumnMapping> newMappings = new ArrayList<>();

        for (ExcelColumnOption option : excelColumnOptions) {
            Optional<ExcelColumnMapping> opt = excelColumnMappings.stream().filter(e -> Objects.equals(option.getHeadName(), e.getHeadName())).findFirst();
            if (!opt.isPresent()) {
                throw new ExcelException(String.format("不支持的字段[%s]", option.getHeadName()));
            }
            ExcelColumnMapping mapping = opt.get();
            if (option.getRequired() != null) {
                mapping.required(option.getRequired());
            }
            newMappings.add(mapping);
        }
        this.excelColumnMappings = newMappings;
        return this;
    }

    public ExcelEntity<E> build() {
        return this;
    }

    @Getter
    public static class ExcelColumnMapping<E, Ex, Im> {

        //        transient Integer columnIndex;
        transient ExcelEntity<E> parent;

        private String headName;
        private String dict;
        private boolean dictUsedValue = false;

        private String comment;

        // export
        private ExcelColumnExportMapping<E, Ex> excelColumnExportMapping;

        // import
        private ExcelColumnImportMapping<E, Im> excelColumnImportMapping;

        public ExcelColumnMapping<E, Ex, Im> headName(String headName) {
            this.headName = headName;
            return this;
        }

        public ExcelColumnMapping<E, Ex, Im> dict(String dict) {
            this.dict = dict;
            return this;
        }

        public ExcelColumnMapping<E, Ex, Im> usedValue() {
            this.dictUsedValue = true;
            return this;
        }

        public ExcelColumnMapping<E, Ex, Im> usedLabel() {
            this.dictUsedValue = false;
            return this;
        }

        public ExcelColumnMapping<E, Ex, Im> comment(String comment) {
            this.comment = comment;
            return this;
        }

        public ExcelColumnMapping<E, Ex, Im> jumpNull() {
            this.excelColumnExportMapping.jumpNull(true);
            this.excelColumnImportMapping.jumpNull(true);
            return this;
        }

        public ExcelColumnMapping<E, Ex, Im> jumpEmpty() {
            this.excelColumnExportMapping.jumpEmpty(true);
            this.excelColumnImportMapping.jumpEmpty(true);
            return this;
        }

        public ExcelColumnMapping<E, Ex, Im> autoTrim() {
            this.excelColumnExportMapping.autoTrim(true);
            this.excelColumnImportMapping.autoTrim(true);
            return this;
        }


        public ExcelColumnMapping<E, Ex, Im> required() {
            this.excelColumnImportMapping.required(true);
            return this;
        }

        public ExcelColumnMapping<E, Ex, Im> required(boolean required) {
            this.excelColumnImportMapping.required(required);
            return this;
        }

        public ExcelColumnMapping<E, Ex, Im> getting(Function<E, Ex> getting) {
            this.excelColumnExportMapping.getting(getting);
            return this;
        }

        public ExcelColumnMapping<E, Ex, Im> getting(Function<E, Ex> getting, Function<Ex, String> formatter) {
            this.excelColumnExportMapping.getting(getting);
            this.excelColumnExportMapping.formatter(formatter);
            return this;
        }

        public ExcelColumnMapping<E, Ex, Im> setting(BiConsumer<E, Im> setting) {
            this.excelColumnImportMapping.setting(setting);
            return this;
        }

        public ExcelColumnMapping<E, Ex, Im> setting(BiConsumer<E, Im> setting, Function<String, Im> formatter) {
            this.excelColumnImportMapping.setting(setting);
            this.excelColumnImportMapping.formatter(formatter);
            return this;
        }

        public ExcelColumnMapping<E, String, String> mapping(String headName) {
            return parent.mapping(headName);
        }

        public <T> ExcelColumnMapping<E, T, T> mapping(String headName, Class<T> type) {
            return parent.mapping(headName, type);
        }

        public <Ext, Imt> ExcelColumnMapping<E, Ext, Imt> mapping(String headName, Class<Ext> exportType, Class<Imt> importType) {
            return parent.mapping(headName, exportType, importType);
        }

        public <T> ExcelColumnMapping<E, T, T> type(Class<T> type) {
            ExcelColumnMapping last = this.parent.excelColumnMappings.remove(this.parent.excelColumnMappings.size() - 1);
            return this.parent.mapping(last.getHeadName(), type);
        }

        public <T> ExcelColumnExportMapping<E, T> exportType(Class<T> type) {
            ExcelColumnMapping<E, T, Im> mapping = new ExcelColumnMapping<>();
            mapping.parent = this.parent;
            mapping.headName = this.headName;
            mapping.dict = this.dict;
            mapping.dictUsedValue = this.dictUsedValue;
            mapping.comment = this.comment;

            mapping.excelColumnExportMapping = new ExcelColumnExportMapping<>();
            mapping.excelColumnExportMapping.parent = mapping;
            mapping.excelColumnExportMapping.type(type);
            mapping.excelColumnExportMapping.jumpNull(this.excelColumnExportMapping.isJumpNull());
            mapping.excelColumnExportMapping.jumpEmpty(this.excelColumnExportMapping.isJumpEmpty());

            mapping.excelColumnImportMapping = this.excelColumnImportMapping;
            mapping.excelColumnImportMapping.parent = mapping;

            mapping.parent.excelColumnMappings.set(mapping.parent.excelColumnMappings.size() - 1, mapping);
            return mapping.excelColumnExportMapping;
        }

        public <T> ExcelColumnImportMapping<E, T> importType(Class<T> type) {
            ExcelColumnMapping<E, Ex, T> mapping = new ExcelColumnMapping<>();
            mapping.parent = this.parent;
            mapping.headName = this.headName;
            mapping.dict = this.dict;
            mapping.dictUsedValue = this.dictUsedValue;
            mapping.comment = this.comment;

            mapping.excelColumnImportMapping = new ExcelColumnImportMapping<>();
            mapping.excelColumnImportMapping.parent = mapping;
            mapping.excelColumnImportMapping.type(type);
            mapping.excelColumnImportMapping.jumpNull(this.excelColumnImportMapping.isJumpNull());
            mapping.excelColumnImportMapping.jumpEmpty(this.excelColumnImportMapping.isJumpEmpty());
            mapping.excelColumnImportMapping.required(this.excelColumnImportMapping.isRequired());

            mapping.excelColumnExportMapping = this.excelColumnExportMapping;
            mapping.excelColumnExportMapping.parent = mapping;

            mapping.parent.excelColumnMappings.set(mapping.parent.excelColumnMappings.size() - 1, mapping);
            return mapping.excelColumnImportMapping;
        }

        public ExcelEntity<E> build() {
            return this.parent.build();
        }

        public ExcelEntity<E> horizontalCellStyleStrategy(HorizontalCellStyleStrategy horizontalCellStyleStrategy) {
            return this.parent.horizontalCellStyleStrategy(horizontalCellStyleStrategy);
        }
    }

    @Getter
    public static class ColumnMapping<E, T> {
        transient ExcelColumnMapping<E, ?, ?> parent;
        private Class<T> type;
        private boolean jumpNull = false;
        private boolean jumpEmpty = false;
        private boolean autoTrim = false;

        public ColumnMapping<E, T> type(Class<T> type) {
            this.type = type;
            return this;
        }

        public ColumnMapping<E, T> jumpNull(boolean jumpNull) {
            this.jumpNull = jumpNull;
            return this;
        }

        public ColumnMapping<E, T> jumpEmpty(boolean jumpEmpty) {
            this.jumpEmpty = jumpEmpty;
            return this;
        }

        public ColumnMapping<E, T> autoTrim() {
            this.autoTrim = true;
            return this;
        }

        public ColumnMapping<E, T> autoTrim(boolean autoTrim) {
            this.autoTrim = autoTrim;
            return this;
        }

        public ExcelColumnMapping<E, ?, ?> headName(String headName) {
            return this.parent.headName(headName);
        }

        public ExcelColumnMapping<E, ?, ?> dict(String dict) {
            return this.parent.dict(dict);
        }

        public ExcelColumnMapping<E, ?, ?> usedValue() {
            return this.parent.usedValue();
        }

        public ExcelColumnMapping<E, ?, ?> usedLabel() {
            return this.parent.usedLabel();
        }

        public ExcelColumnMapping<E, ?, ?> comment(String comment) {
            return this.parent.comment(comment);
        }

        public ExcelColumnMapping<E, String, String> mapping(String headName) {
            return parent.mapping(headName);
        }

        public <TT> ExcelColumnMapping<E, TT, TT> mapping(String headName, Class<TT> type) {
            return parent.mapping(headName, type);
        }

        public <Ext, Imt> ExcelColumnMapping<E, Ext, Imt> mapping(String headName, Class<Ext> exportType, Class<Imt> importType) {
            return parent.mapping(headName, exportType, importType);
        }

        public ExcelEntity<E> build() {
            return this.parent.build();
        }

        public ExcelEntity<E> horizontalCellStyleStrategy(HorizontalCellStyleStrategy horizontalCellStyleStrategy) {
            return this.parent.horizontalCellStyleStrategy(horizontalCellStyleStrategy);
        }
    }

    @Getter
    @Setter
    public static class ExcelColumnExportMapping<E, Ex> extends ColumnMapping<E, Ex> {
        private Function<E, Ex> getting;
        private Function<Ex, String> formatter;

        public ExcelColumnExportMapping<E, Ex> getting(Function<E, Ex> getting) {
            this.getting = getting;
            return this;
        }

        public <T> ExcelColumnImportMapping<E, T> importType(Class<T> type) {
            return parent.importType(type);
        }

        public ExcelColumnExportMapping<E, Ex> formatter(Function<Ex, String> formatter) {
            this.formatter = formatter;
            return this;
        }
    }

    @Getter
    @Setter
    public static class ExcelColumnImportMapping<E, Im> extends ColumnMapping<E, Im> {
        private BiConsumer<E, Im> setting;
        private boolean required;
        private Function<String, Im> formatter;

        public ExcelColumnImportMapping<E, Im> setting(BiConsumer<E, Im> setting) {
            this.setting = setting;
            return this;
        }

        public ExcelColumnImportMapping<E, Im> formatter(Function<String, Im> formatter) {
            this.formatter = formatter;
            return this;
        }

        public ExcelColumnImportMapping<E, Im> required(boolean required) {
            this.required = required;
            return this;
        }

        public <T> ExcelColumnExportMapping<E, T> exportType(Class<T> type) {
            return parent.exportType(type);
        }
    }

}
