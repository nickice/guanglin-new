package com.guanglin.pptGen.model.pptConfig;

import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;

@Builder
@Data
@AllArgsConstructor
public class Font {
    private String family;
    private String size;
    private String color;
    private String style;
    private String content;

}
