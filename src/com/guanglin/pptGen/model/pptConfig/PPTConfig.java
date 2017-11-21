package com.guanglin.pptGen.model.pptConfig;


import lombok.Builder;
import lombok.Data;

@Data
@Builder
public class PPTConfig {

    public SlideConfig firstSlideConfig;
    public SlideConfig otherSlideConfig;
    public SlideConfig lastSlideConfig;

}
