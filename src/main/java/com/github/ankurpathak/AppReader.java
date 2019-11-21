package com.github.ankurpathak;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.jxls.reader.ReaderBuilder;
import org.jxls.reader.XLSReader;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.xml.sax.SAXException;

import java.io.*;
import java.util.*;

/**
 * Hello world!
 *
 */
public class AppReader
{
    static Logger logger = LoggerFactory.getLogger(XlsReaderDemo.class);
    private static String dataFile = "reader.xls";
    public static final String xmlConfig = "reader.xml";

    public static void main(String[] args) throws IOException, SAXException, InvalidFormatException {
        logger.info("Running XLS Reader demo");
        execute();
    }

    public static void execute() throws IOException, SAXException, InvalidFormatException {
        logger.info("Reading xml config file and constructing XLSReader");
        List<YearDto> items = new ArrayList<YearDto>();
        try(InputStream xmlInputStream = AppReader.class.getResourceAsStream(xmlConfig)) {
            XLSReader reader = ReaderBuilder.buildFromXML(xmlInputStream);
            try (InputStream xlsInputStream = AppReader.class.getResourceAsStream(dataFile)) {
                Map<String, Object> beans = new HashMap<>();
                YearDto item = new YearDto();
                 beans.put("item", item);
                 beans.put("items",items);
                 logger.info("Reading the data...");
                 reader.read(xlsInputStream, beans);
                 logger.info("Read " +  item + " item");
            }
        }
    }


}
