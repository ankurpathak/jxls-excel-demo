package com.github.ankurpathak;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.jxls.reader.ReaderBuilder;
import org.jxls.reader.XLSReader;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.xml.sax.SAXException;

import java.io.IOException;
import java.io.InputStream;
import java.util.HashMap;
import java.util.Map;

/**
 * Hello world!
 *
 */
public class ProfitLossStatementReaderUtil
{

    static Logger logger = LoggerFactory.getLogger(ProfitLossStatementReaderUtil.class);
    private static String dataFile = "financial-reader.xls";
    public static final String xmlConfig = "financial-reader.xml";

    public static void main(String[] args) throws IOException, SAXException, InvalidFormatException {
        logger.info("Running XLS Reader demo");
        execute();
    }

    public static Map<String, Map<String, String>> execute() throws IOException, SAXException, InvalidFormatException {
        logger.info("Reading xml config file and constructing XLSReader");
        try(InputStream xmlInputStream = ProfitLossStatementReaderUtil.class.getResourceAsStream(xmlConfig)) {
            XLSReader reader = ReaderBuilder.buildFromXML(xmlInputStream);
            try (InputStream xlsInputStream = ProfitLossStatementReaderUtil.class.getResourceAsStream(dataFile)) {
                Map<String, Object> beans = new HashMap<>();
                YearDto yearDto =     new YearDto();
                beans.put("item", yearDto);
                logger.info("Reading the data...");
                reader.read(xlsInputStream, beans);
                logger.info("Read ");

            }
        }
        return null;
    }



    public static class YearDto {
            String firstYear;
            String secondYear;
            String thirdYear;
            String fourthYear;
            String fifthYear;

        public String getFirstYear() {
            return firstYear;
        }

        public void setFirstYear(String firstYear) {
            this.firstYear = firstYear;
        }

        public String getSecondYear() {
            return secondYear;
        }

        public void setSecondYear(String secondYear) {
            this.secondYear = secondYear;
        }

        public String getThirdYear() {
            return thirdYear;
        }

        public void setThirdYear(String thirdYear) {
            this.thirdYear = thirdYear;
        }

        public String getFourthYear() {
            return fourthYear;
        }

        public void setFourthYear(String fourthYear) {
            this.fourthYear = fourthYear;
        }

        public String getFifthYear() {
            return fifthYear;
        }

        public void setFifthYear(String fifthYear) {
            this.fifthYear = fifthYear;
        }
    }



}
