package com.guanglin.pptGen.model.pptConfig;

import lombok.Builder;
import lombok.Data;

@Builder
@Data
public class Font {
    public String family;
    public String size;
    public String color;
    public String style;

}
