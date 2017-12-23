package com.guanglin.pptGen.project;

import com.guanglin.pptGen.exception.ProjectException;
import com.guanglin.pptGen.model.Capture;
import com.guanglin.pptGen.model.Item;
import com.guanglin.pptGen.model.Project;

import java.io.File;
import java.util.Stack;

/**
 * Created by pengyao on 02/07/2017.
 */
public class ProjectBuilder {


    /*
    public static Project builder() {

    }
    */

    private static Project mapCaptures(Project project, String captureFolderPath) throws ProjectException {

        if (project == null || project.getItems() == null || project.getItems().size() == 0) {
            throw new ProjectException("项目数据是null。");
        }

        try {
            File file = new File(captureFolderPath);
            if (file.isDirectory()) {

                for (Item item : project.getItems()) {
                    String itemCaptureFolderPath = captureFolderPath + item.getFields().get("项目名称");
                    File captureFolder = new File(itemCaptureFolderPath);
                    if (captureFolder == null || !captureFolder.isDirectory()) {
                        continue;
                    }

                    Stack<Capture> insideCaptures = new Stack<Capture>();
                    for (File f : captureFolder.listFiles()) {

                        Capture capture = new Capture(f.getName(), f.getAbsolutePath());

                        if (f.getName().contains("outdoor")) {
                            item.setOutsideCapture(capture);
                        } else {
                            insideCaptures.push(capture);
                        }
                    }

                    item.setInsideCaptures(insideCaptures);
                }
            }
        } catch (Exception ex) {
            throw new ProjectException(ex.getMessage());

        }

        return project;
    }

}
