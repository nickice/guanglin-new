package com.guanglin.pptGen.utils;

import com.fasterxml.jackson.databind.ObjectMapper;

import java.io.File;
import java.io.IOException;


public class FormatCoverter {

    private static ObjectMapper objectMapper = new ObjectMapper();

    public static <T> T convertJsonToObject(String jsonFilePath, Class<T> clazz) throws IOException {
        return objectMapper.readValue(new File(jsonFilePath), clazz);
    }
}
