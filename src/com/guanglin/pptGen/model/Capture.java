package com.guanglin.pptGen.model;

import com.google.common.base.Strings;
import lombok.Data;

import java.io.FileInputStream;
import java.io.IOException;

/**
 * Created by pengyao on 01/06/2017.
 */
public @Data
class Capture {
    private String captureName;
    private String capturePath;

    public Capture(String captureName, String capturePath) {
        this.captureName = captureName;
        this.capturePath = capturePath;
    }


    public byte[] getBytes() {
        FileInputStream captureStream = null;

        try {

            // no image path, then return null
            if(Strings.isNullOrEmpty(this.capturePath)) {
                return null;
            }

            captureStream = new FileInputStream(this.capturePath);

            byte[] bytes = new byte[captureStream.available()];
            captureStream.read(bytes);
            return bytes;
        } catch (IOException e) {
            return null;
        }
        finally {
            if(captureStream != null) {
                try {
                    captureStream.close();
                } catch (IOException e) {
                    return null;
                }
            }
        }
    }
}
