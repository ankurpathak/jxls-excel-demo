package com.github.ankurpathak;

import org.jxls.area.Area;
import org.jxls.builder.AreaBuilder;
import org.jxls.builder.xml.XmlAreaBuilder;
import org.jxls.common.CellRef;
import org.jxls.common.Context;
import org.jxls.transform.Transformer;
import org.jxls.util.TransformerFactory;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.*;

/**
 * Hello world!
 *
 */
public class App 
{

    static Logger logger = LoggerFactory.getLogger(App.class);
    private static String template = "demo.xls";
    private static String xmlConfig = "demo.xml";
    private static String output = "target/demo_output.xls";
    public static void main( String[] args ) throws Exception
    {
        logger.info("Opening input stream");
        try(InputStream is = App.class.getResourceAsStream(template)) {
            try (OutputStream os = new FileOutputStream(output)) {
                Transformer transformer = TransformerFactory.createTransformer(is, os);
                System.out.println("Creating areas");
                InputStream configInputStream = App.class.getResourceAsStream(xmlConfig);
                AreaBuilder areaBuilder = new XmlAreaBuilder(configInputStream, transformer);
                List<Area> xlsAreaList = areaBuilder.build();
                Area xlsArea = xlsAreaList.get(0);
                Context context = new Context();
                List<Map<String, String>> names = new ArrayList<>();
                Map<String, String> nameOne = new HashMap<>();
                nameOne.put("firstName", "Ankur");
                nameOne.put("lastName", "Pathak");
                names.add(nameOne);
                Map<String, String> nameTwo = new HashMap<>();
                nameTwo.put("firstName", "Mamta");
                nameTwo.put("lastName", "Pathak");
                names.add(nameTwo);

                /*List<Name> names =   new ArrayList<>();
                names.addAll(Arrays.asList(new Name("Ankur", "Pathak"), new Name("Mamta", "Pathak")));*/
                context.putVar("names", names);
                logger.info("Applying first area at cell " + new CellRef("Template!A1"));
                xlsArea.applyAt(new CellRef("Template!A1"), context);
                xlsArea.processFormulas();
                logger.info("Complete");
                transformer.write();
                logger.info("written to file");
            }
        }
    }
}
