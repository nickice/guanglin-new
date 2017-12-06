package com.guanglin.pptGen.model.pptConfig;

import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;

@Data
@Builder
@AllArgsConstructor
public class Image {
    private String height;
    private String width;
    private String size;
}
