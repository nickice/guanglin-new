package com.guanglin.pptGen.utils;

import com.google.common.base.Strings;

import java.io.FileInputStream;
import java.io.IOException;

public class ImageUtils {
    public static byte[] getImageBytes(final String imagePath) {
        FileInputStream captureStream = null;

        try {

            // no image path, then return null
            if (Strings.isNullOrEmpty(imagePath)) {
                return null;
            }

            captureStream = new FileInputStream(imagePath);

            byte[] bytes = new byte[captureStream.available()];
            captureStream.read(bytes);
            return bytes;
        } catch (IOException e) {
            return null;
        } finally {
            if (captureStream != null) {
                try {
                    captureStream.close();
                } catch (IOException e) {
                    return null;
                }
            }
        }
    }
}
