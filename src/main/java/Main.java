import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.xmlbeans.XmlObject;
import org.apache.xmlbeans.XmlString;
import org.json.simple.JSONArray;
import org.json.simple.JSONObject;
import org.openxmlformats.schemas.presentationml.x2006.main.CTSlide;

import java.io.*;

public class Main {


    public static void main(String[] args) throws IOException {

        JSONArray ppts = new JSONArray();
        File dir = new File("input/");
        File[] directoryListing = dir.listFiles();
        if (directoryListing != null) {
            for (File child : directoryListing) {
                FileInputStream input = new FileInputStream(child);
                String filename = child.getName();

                JSONObject slideSetObject = new JSONObject();
                JSONArray arr = new JSONArray();
                XMLSlideShow slideShow = new XMLSlideShow(input);
                slideSetObject.put("slideSet", filename);
                for (XSLFSlide slide : slideShow.getSlides()) {
                    JSONObject obj = new JSONObject();
                    obj.put("slide", slide.getSlideNumber());
                    CTSlide ctSlide = slide.getXmlObject();
                    XmlObject[] allText = ctSlide.selectPath(
                            "declare namespace a='http://schemas.openxmlformats.org/drawingml/2006/main' " + ".//a:t");
                    StringBuilder text = new StringBuilder();
                    for (XmlObject xmlObject : allText) {
                        if (xmlObject instanceof XmlString) {
                            XmlString xmlString = (XmlString) xmlObject;
                            text.append(" ").append(xmlString.getStringValue());

                        }
                    }
                    obj.put("text", text.toString());
                    arr.add(obj);
                }
                slideSetObject.put("slides",arr);
                ppts.add(slideSetObject);
            }
        }

        try (FileWriter file = new FileWriter("output/pptToText.json")) {
            file.write(ppts.toJSONString());
            System.out.println("Successfully Copied JSON Object to File...");
            System.out.println("\nJSON Object: " + ppts);
        }

    }

}