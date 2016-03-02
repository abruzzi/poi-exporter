package com.thoughtworks.ppt.exporter.demo;

import org.apache.poi.hslf.HSLFSlideShow;
import org.apache.poi.hslf.model.AutoShape;
import org.apache.poi.hslf.model.Freeform;
import org.apache.poi.hslf.model.Slide;
import org.apache.poi.hslf.usermodel.SlideShow;
import org.apache.poi.ss.usermodel.ShapeTypes;

import java.io.FileOutputStream;
import java.io.IOException;

public class Main {
    public static void main(String[] args) {
        HSLFSlideShow ppt = HSLFSlideShow.create();
        SlideShow slideShow = new SlideShow(ppt);
        Slide slide = slideShow.createSlide();

//        Line line = new Line();
//        line.setAnchor(new java.awt.Rectangle(50, 50, 100, 2));
//        line.setShapeType(ShapeTypes.LINE);
//        line.setLineColor(Color.black);
//        line.setLineStyle(HSSFShape.LINESTYLE_DASHDOTDOTSYS);
//
//        slide.addShape(line);

        AutoShape node1 = new AutoShape(ShapeTypes.STAR_4);
        node1.setAnchor(new java.awt.Rectangle(50, 50, 100, 100));

        AutoShape node2 = new AutoShape(ShapeTypes.STAR_4);
        node2.setAnchor(new java.awt.Rectangle(150, 150, 100, 100));

        slide.addShape(node1);
        slide.addShape(node2);

        java.awt.geom.GeneralPath path = new java.awt.geom.GeneralPath();
        path.moveTo(100, 100);
        path.lineTo(200, 100);
        path.lineTo(200, 200);
        path.lineTo(400, 200);
        path.lineTo(400, 250);
        path.lineTo(100, 250);

        AutoShape line = new AutoShape(ShapeTypes.LINE);
        line.setAnchor(new java.awt.Rectangle(300, 300, 100, 1));
        slide.addShape(line);

        path.closePath();

        Freeform shape = new Freeform();
        shape.setPath(path);

        slide.addShape(shape);

        try {
            ppt.write(new FileOutputStream("slides.ppt"));
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
