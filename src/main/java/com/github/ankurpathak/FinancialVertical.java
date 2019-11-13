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

import static com.github.ankurpathak.FinancialVertical.CAMKeys.*;

/**
 * Hello world!
 *
 *
 */
public class FinancialVertical
{

    static Logger logger = LoggerFactory.getLogger(FinancialVertical.class);
    private static final String template = "financial.xls";
    private static final String xmlConfig = "financial.xml";
    private static final String output = "target/financial_output.xls";
    public static void main( String[] args ) throws Exception
    {
        logger.info("Opening input stream");
        try(InputStream is = FinancialVertical.class.getResourceAsStream(template)) {
            try (OutputStream os = new FileOutputStream(output)) {
                Transformer transformer = TransformerFactory.createTransformer(is, os);
                System.out.println("Creating areas");
                InputStream configInputStream = FinancialVertical.class.getResourceAsStream(xmlConfig);
                AreaBuilder areaBuilder = new XmlAreaBuilder(configInputStream, transformer);
                List<Area> xlsAreaList = areaBuilder.build();
                Area xlsArea = xlsAreaList.get(0);
                Context context = new Context();
                List<Map<String, String>> aFinancials = new ArrayList<>();
                aFinancials.add(Collections.emptyMap());
                aFinancials.add(Collections.emptyMap());
                aFinancials.add(Collections.emptyMap());


                List<Map<String, String>> aGrowths = new ArrayList<>();
                aGrowths.add(Collections.emptyMap());
                aGrowths.add(Collections.emptyMap());


                context.putVar("aFinancials", aFinancials);
                context.putVar("aGrowths", aGrowths);
                logger.info("Applying first area at cell " + new CellRef("Template!A1"));
                xlsArea.applyAt(new CellRef("Template!A1"), context);
                xlsArea.processFormulas();
                logger.info("Complete");
                transformer.write();
                logger.info("written to file");
            }
        }
    }


    public interface  CAMKeys {
        String aFinancial = "aFinancial";
        String aYears = "aYears";
        String oSales = "oSales";
        String oExciseDuty = "oExciseDuty";
        String oNetSales = "oNetSales";
        String oExportIncentives = "oExportIncentives";
        String oTotalOperatingIncome = "oTotalOperatingIncome";
        String oRawMaterialsImported = "oRawMaterialsImported";
        String oOtherMfgExpenses = "oOtherMfgExpenses";
        String oRepairsAndMaintenance = "oRepairsAndMaintenance";
        String oCostOfGoodsSold = "oCostOfGoodsSold";
        String oGrossProfits = "oGrossProfits";
        String oSellingGeneralAdmExpenses = "oSellingGeneralAdmExpenses";
        String oMiscExpensesWrittenOff = "oMiscExpensesWrittenOff";
        String oTransportOrCartage = "oTransportOrCartage";
        String oOpbdit = "oOpbdit";
        String oOpbit = "oOpbit";
        String oInterestAndOtherFinanceCharges = "oInterestAndOtherFinanceCharges";
        String oOpbt = "oOpbt";
        String oTotalNonOperatingIncome = "oTotalNonOperatingIncome";
        String oTotalNonOperatingExpense = "oTotalNonOperatingExpense";
        String oPBT = "oPBT";
        String oTaxPaid = "oTaxPaid";
        String oPAT = "oPAT";
        String oAdjPat = "oAdjPat";
        String oEquityDividendPaid = "oEquityDividendPaid";
        String oPartnersWithDrawl = "oPartnersWithDrawl";
        String oDividendTax = "oDividendTax";
        String oDividendPreference = "oDividendPreference";
        String oRetainedProfit = "oRetainedProfit";
        String oShareCapitalPaidUp = "oShareCapitalPaidUp";
        String oSharePremiumAcc = "oSharePremiumAcc";
        String oTotalShareCapital = "oTotalShareCapital";
        String oRevaluationReserve = "oRevaluationReserve";
        String oOtherReserves = "oOtherReserves";
        String oMiscExpNotWrittenOff = "oMiscExpNotWrittenOff";
        String oTotalNetWorth = "oTotalNetWorth";
        String oBorrowingsFromAffiliatesAssociates = "oBorrowingsFromAffiliatesAssociates";
        String oShortTermBorrowingsFromOthers = "oShortTermBorrowingsFromOthers";
        String oOtherTermLiabilities = "oOtherTermLiabilities";
        String oBorrowingsFromSubsidiaries = "oBorrowingsFromSubsidiaries";
        String oDeferedTaxLiability = "oDeferedTaxLiability";
        String oLongTermProvision = "oLongTermProvision";
        String oTotalTermLiabilities = "oTotalTermLiabilities";
        String oShortTermBorrowingsFromAssociates = "oShortTermBorrowingsFromAssociates";
        String oCreditorsandBillsPayable = "oCreditorsandBillsPayable";
        String oWorkingCapitalLimitfromBanks = "oWorkingCapitalLimitfromBanks";
        String oTotalRepaymentsDueWithinYear = "oTotalRepaymentsDueWithinYear";
        String oOtherCurrLiabilitiesAndProvisions = "oOtherCurrLiabilitiesAndProvisions";
        String oProvisionForTaxation = "oProvisionForTaxation";
        String oOthercurrentLiability1 = "oOthercurrentLiability1";
        String oOthercurrentLiability2 = "oOthercurrentLiability2";
        String oOthercurrentLiability3 = "oOthercurrentLiability3";
        String oCurrentLiabilities = "oCurrentLiabilities";
        String oTotalLiabilities = "oTotalLiabilities";
        String oBalanceSheetTotal = "oBalanceSheetTotal";
        String oGrossBlock = "oGrossBlock";
        String oAccumulatedDepreciation = "oAccumulatedDepreciation";
        String oNetBlock = "oNetBlock";
        String oCapitalWip = "oCapitalWip";
        String oDeferredRevenueExpenditure = "oDeferredRevenueExpenditure";
        String oGovtAndOtherSecurities = "oGovtAndOtherSecurities";
        String oTotalOtherNonCurrentAssets = "oTotalOtherNonCurrentAssets";
        String oDebtorsMoreThanSixMonths = "oDebtorsMoreThanSixMonths";
        String oLoansAndAdvancess = "oLoansAndAdvancess";
        String oFixedDepositsWithBanks = "oFixedDepositsWithBanks";
        String oOtherNonCurrentAssets = "oOtherNonCurrentAssets";
        String oDeferredTaxAsset = "oDeferredTaxAsset";
        String oTotalCurrentAssets = "oTotalCurrentAssets";
        String oTotalInventory = "oTotalInventory";
        String oReceivablesOtherThanDeferred = "oReceivablesOtherThanDeferred";
        String oDeferredReceivables = "oDeferredReceivables";
        String oCashAndBankBalances = "oCashAndBankBalances";
        String oOtherCurrentAssets = "oOtherCurrentAssets";
        String oAdvancesToSuppliers = "oAdvancesToSuppliers";
        String oOthercurrentAsset1 = "oOthercurrentAsset1";
        String oOthercurrentAsset2 = "oOthercurrentAsset2";
        String oOthercurrentAsset3 = "oOthercurrentAsset3";
        String oTotalAssets = "oTotalAssets";
        String oDifferenceFields = "oDifferenceFields";
        String TOLTNW = "TOLTNW";
        String RETURNONCAPITALEMPLOYED = "RETURNONCAPITALEMPLOYED";
        String EBIDTABYSALES = "EBIDTABYSALES";
        String EBIDTAINTEREST = "EBIDTAINTEREST";
        String ASSETTURNOVERRATIO = "ASSETTURNOVERRATIO";
        String CREDITORSTURNOVER = "CREDITORSTURNOVER" ;
        String NETWORKINGCAPITAL = "NETWORKINGCAPITAL";
        String AVGCOLLECTIONPERIOD = "AVGCOLLECTIONPERIOD" ;
        String AVERAGEDAYS = "AVERAGEDAYS";
        String INVENTORYTOCOSTOFGOODSSOLD = "INVENTORYTOCOSTOFGOODSSOLD";
        String CURRENTRATIO = "CURRENTRATIO";
        String LIQUIDITYRATIO = "LIQUIDITYRATIO";
        String LONGDEBTTOEQUITY = "LONGDEBTTOEQUITY";
        String EQUITY = "EQUITY";
        String INTERSETCOVERAGERATIO = "INTERSETCOVERAGERATIO";
        String DSCR = "DSCR";
        String GROSSMARGINS = "GROSSMARGINS";
        String NETPROFITRATIO = "NETPROFITRATIO";
        String CASHPROFITBYRATIO = "CASHPROFITBYRATIO";
        String GROWTHINSALES = "GROWTHINSALES";
        String GROWTHINPAT = "GROWTHINPAT";
        String CASHACCRUALS = "CASHACCRUALS";
        String oNetProfitAfterTax = "oNetProfitAfterTax";
        String oDepreciation = "oDepreciation";
        String oSalaryPaidPromPartner = "oSalaryPaidPromPartner";
        String oInterestPaidPromoters = "oInterestPaidPromoters";
        String oProvisionTax = "oProvisionTax";
        String oInterest = "oInterest";
        String oOtherIncome = "oOtherIncome";
        String oOperatingCashProfit = "oOperatingCashProfit";
    }

    public static Map<String, String> prepareFinancial(){
        Map<String, String> aFinancial = new HashMap<>();
        aFinancial.put(CAMKeys.aFinancial, "");
        aFinancial.put(CAMKeys.aYears, "");
        aFinancial.put(CAMKeys.oSales, "");
        aFinancial.put(CAMKeys.oExciseDuty, "");
        aFinancial.put(CAMKeys.oNetSales, "");
        aFinancial.put(CAMKeys.oExportIncentives, "");
        aFinancial.put(CAMKeys.oTotalOperatingIncome, "");

        aFinancial.put(CAMKeys.oRawMaterialsImported, "");
        aFinancial.put(CAMKeys.oOtherMfgExpenses, "");
        aFinancial.put(CAMKeys.oRepairsAndMaintenance, "");
        aFinancial.put(CAMKeys.oCostOfGoodsSold, "");
        aFinancial.put(CAMKeys.oCostOfGoodsSold, "");
        aFinancial.put(CAMKeys.oGrossProfits, "");
        aFinancial.put(CAMKeys.oSellingGeneralAdmExpenses, "");
        aFinancial.put(CAMKeys.oMiscExpensesWrittenOff, "");
        aFinancial.put(CAMKeys.oTransportOrCartage, "");
        aFinancial.put(CAMKeys.oOpbdit, "");
        aFinancial.put(CAMKeys.oOpbit, "");
        aFinancial.put(CAMKeys.oInterestAndOtherFinanceCharges, "");
        aFinancial.put(CAMKeys.oOpbt, "");
        aFinancial.put(CAMKeys.oTotalNonOperatingIncome, "");
        aFinancial.put(CAMKeys.oTotalNonOperatingExpense, "");
        aFinancial.put(CAMKeys.oPBT, "");
        aFinancial.put(CAMKeys.oTaxPaid, "");
        aFinancial.put(CAMKeys.oPAT, "");
        aFinancial.put(CAMKeys.oAdjPat, "");
        aFinancial.put(CAMKeys.oEquityDividendPaid, "");
        aFinancial.put(CAMKeys.oPartnersWithDrawl, "");
        aFinancial.put(CAMKeys.oDividendTax, "");
        aFinancial.put(CAMKeys.oDividendPreference, "");
        aFinancial.put(CAMKeys.oRetainedProfit, "");
        aFinancial.put(CAMKeys.oShareCapitalPaidUp, "");
        aFinancial.put(CAMKeys.oSharePremiumAcc, "");
        aFinancial.put(CAMKeys.oTotalShareCapital, "");
        aFinancial.put(CAMKeys.oRevaluationReserve, "");
        aFinancial.put(CAMKeys.oOtherReserves, "");
        aFinancial.put(CAMKeys.oMiscExpNotWrittenOff, "");
        aFinancial.put(CAMKeys.oTotalNetWorth, "");
        aFinancial.put(CAMKeys.oBorrowingsFromAffiliatesAssociates, "");
        aFinancial.put(CAMKeys.oShortTermBorrowingsFromOthers, "");
        aFinancial.put(CAMKeys.oOtherTermLiabilities, "");
        aFinancial.put(CAMKeys.oBorrowingsFromSubsidiaries, "");
        aFinancial.put(CAMKeys.oDeferedTaxLiability, "");
        aFinancial.put(CAMKeys.oLongTermProvision, "");
        aFinancial.put(CAMKeys.oTotalTermLiabilities, "");
        aFinancial.put(CAMKeys.oShortTermBorrowingsFromAssociates, "");
        aFinancial.put(CAMKeys.oCreditorsandBillsPayable, "");
        aFinancial.put(CAMKeys.oWorkingCapitalLimitfromBanks, "");
        aFinancial.put(CAMKeys.oTotalRepaymentsDueWithinYear, "");
        aFinancial.put(CAMKeys.oOtherCurrLiabilitiesAndProvisions, "");
        aFinancial.put(CAMKeys.oProvisionForTaxation, "");
        aFinancial.put(CAMKeys.oOthercurrentLiability1, "");
        aFinancial.put(CAMKeys.oOthercurrentLiability2, "");
        aFinancial.put(CAMKeys.oOthercurrentLiability3, "");
        aFinancial.put(CAMKeys.oCurrentLiabilities, "");
        aFinancial.put(CAMKeys.oTotalLiabilities, "");
        aFinancial.put(CAMKeys.oBalanceSheetTotal, "");
        aFinancial.put(CAMKeys.oGrossBlock, "");
        aFinancial.put(CAMKeys.oAccumulatedDepreciation, "");
        aFinancial.put(CAMKeys.oNetBlock, "");
        aFinancial.put(CAMKeys.oCapitalWip, "");
        aFinancial.put(CAMKeys.oDeferredRevenueExpenditure, "");
        aFinancial.put(CAMKeys.oGovtAndOtherSecurities, "");
        aFinancial.put(CAMKeys.oTotalOtherNonCurrentAssets, "");
        aFinancial.put(CAMKeys.oDebtorsMoreThanSixMonths, "");
        aFinancial.put(CAMKeys.oLoansAndAdvancess, "");
        aFinancial.put(CAMKeys.oFixedDepositsWithBanks, "");
        aFinancial.put(CAMKeys.oOtherNonCurrentAssets, "");
        aFinancial.put(CAMKeys.oDeferredTaxAsset, "");
        aFinancial.put(CAMKeys.oTotalCurrentAssets, "");
        aFinancial.put(CAMKeys.oTotalInventory, "");
        aFinancial.put(CAMKeys.oReceivablesOtherThanDeferred, "");
        aFinancial.put(CAMKeys.oDeferredReceivables, "");
        aFinancial.put(CAMKeys.oCashAndBankBalances, "");
        aFinancial.put(CAMKeys.oOtherCurrentAssets, "");
        aFinancial.put(CAMKeys.oAdvancesToSuppliers, "");
        aFinancial.put(CAMKeys.oOthercurrentAsset1, "");
        aFinancial.put(CAMKeys.oOthercurrentAsset2, "");
        aFinancial.put(CAMKeys.oOthercurrentAsset3, "");
        aFinancial.put(CAMKeys.oTotalAssets, "");
        aFinancial.put(CAMKeys.oDifferenceFields, "");
        aFinancial.put(CAMKeys.TOLTNW, "");
        aFinancial.put(CAMKeys.RETURNONCAPITALEMPLOYED, "");
        aFinancial.put(CAMKeys.EBIDTABYSALES, "");
        aFinancial.put(CAMKeys.EBIDTAINTEREST, "");
        aFinancial.put(CAMKeys.ASSETTURNOVERRATIO, "");
        aFinancial.put(CAMKeys.CREDITORSTURNOVER, "");
        aFinancial.put(CAMKeys.NETWORKINGCAPITAL, "");
        aFinancial.put(CAMKeys.AVGCOLLECTIONPERIOD, "");
        aFinancial.put(CAMKeys.AVERAGEDAYS, "");
        aFinancial.put(CAMKeys.INVENTORYTOCOSTOFGOODSSOLD, "");
        aFinancial.put(CAMKeys.CURRENTRATIO, "");
        aFinancial.put(CAMKeys.LIQUIDITYRATIO, "");
        aFinancial.put(CAMKeys.LONGDEBTTOEQUITY, "");
        aFinancial.put(CAMKeys.EQUITY, "");
        aFinancial.put(CAMKeys.INTERSETCOVERAGERATIO, "");
        aFinancial.put(CAMKeys.DSCR, "");
        aFinancial.put(CAMKeys.GROSSMARGINS, "");
        aFinancial.put(CAMKeys.NETPROFITRATIO, "");
        aFinancial.put(CAMKeys.CASHPROFITBYRATIO, "");
        aFinancial.put(CAMKeys.GROWTHINSALES, "");
        aFinancial.put(CAMKeys.GROWTHINPAT, "");
        aFinancial.put(CAMKeys.CASHACCRUALS, "");
        aFinancial.put(CAMKeys.oNetProfitAfterTax, "");
        aFinancial.put(CAMKeys.oDepreciation, "");
        aFinancial.put(CAMKeys.oSalaryPaidPromPartner, "");
        aFinancial.put(CAMKeys.oInterestPaidPromoters, "");
        aFinancial.put(CAMKeys.oProvisionTax, "");
        aFinancial.put(CAMKeys.oInterest, "");
        aFinancial.put(CAMKeys.oOtherIncome, "");
        aFinancial.put(CAMKeys.oOperatingCashProfit, "");


        ${oTradeOtherReceivable}
        ${oInventories}
        ${oLoanAdvance}
        ${oOtherCurrentLiabilities}
        ${oWorkingCapitalChanges}
        ${oCashGenratedOperations}
        ${oTaxesPaid}
        ${oNetCashOperations}

    }

}
