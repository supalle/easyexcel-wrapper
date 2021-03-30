package com.supalle.easyexcel.wrapper;

import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;
import lombok.NoArgsConstructor;

import java.util.List;
import java.util.Objects;


@Data
@Builder
@NoArgsConstructor
@AllArgsConstructor
public class Dict {
    private String dictName;
    private String comment;
    private List<DictItem> dictItems;

    @Override
    public boolean equals(Object o) {
        if (this == o) return true;
        if (o == null || getClass() != o.getClass()) return false;

        Dict dict = (Dict) o;

        return Objects.equals(dictName, dict.dictName);
    }

    @Override
    public int hashCode() {
        return dictName != null ? dictName.hashCode() : 0;
    }

    @Data
    @Builder
    @NoArgsConstructor
    @AllArgsConstructor
    public static class DictItem {
        private String value;
        private String label;
        private String comment;

        @Override
        public boolean equals(Object o) {
            if (this == o) return true;
            if (o == null || getClass() != o.getClass()) return false;

            DictItem dictItem = (DictItem) o;

            return Objects.equals(value, dictItem.value);
        }

        @Override
        public int hashCode() {
            return value != null ? value.hashCode() : 0;
        }

        @Override
        public String toString() {
            return value + ":" + label;
        }
    }
}
