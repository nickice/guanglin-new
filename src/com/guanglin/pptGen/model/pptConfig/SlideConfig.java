package com.guanglin.pptGen.model.pptConfig;

import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;

@Builder
@Data
@AllArgsConstructor
public class SlideConfig {
    private Font fontConfig;
    private Image imageConfig;
}
