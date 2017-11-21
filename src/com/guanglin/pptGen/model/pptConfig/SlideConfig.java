package com.guanglin.pptGen.model.pptConfig;


import lombok.Builder;
import lombok.Data;

@Builder
@Data
public class SlideConfig {
    public Font fontConfig;
    public Image imageConfig;
}
