package com.guanglin.pptGen.model.pptConfig;


import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;

@Data
@Builder
@AllArgsConstructor
public class PPTConfig {

    private SlideConfig firstSlideConfig;
    private SlideConfig otherSlideConfig;
    private SlideConfig lastSlideConfig;

}
