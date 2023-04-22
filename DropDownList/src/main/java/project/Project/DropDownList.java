package project.Project;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.File;
import org.apache.poi.ss.usermodel.DataValidation;
import org.apache.poi.ss.usermodel.DataValidationConstraint;
import org.apache.poi.ss.usermodel.DataValidationHelper;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import java.awt.Desktop;
import org.apache.poi.ss.util.*;
import java.util.Map;
import java.util.HashMap;

public class DropDownList {
	public static void main(String[] args) throws Exception {

		// some data
		Map<String, String[]> categoryItems = new HashMap<String, String[]>();
		categoryItems.put("Opening_Balance", new String[] { "OPENING BALANCE AS PER CANDID", "OPENING BALANCE AS PER SRG"});
		categoryItems.put("Credit_Note", new String[] { "DEBIT NOTE ACCOUNTED ON 2021-2021", "DEBIT NOTE NOT ACCOUNTED IN CANDID KNITS LLP", "DEBIT NOTE NOT ACCOUNTED IN SRG BOOK" });
		categoryItems.put("Debit_Note", new String[] { "CREDIT NOTE NOT ACCOUNTED IN CANDID KNITS LLP" });
		categoryItems.put("Auto_Sales", new String[] { "SALES ACCOUNTED IN SRG BOOKS", "SALES NOT ACCOUNTED IN CANDID KNITS LLP" });
		categoryItems.put("Purchase_Manual", new String[] { "PURCHASE ACCOUNTED IN SRG BOOKS" });
		categoryItems.put("Journal", new String[] { "UNKOWN DEBIT IN SRG BOOKS" });
		categoryItems.put("Receipt", new String[] { "RECEIPT ACCOUNTED IN SRG BOOK", "RECEIPT NOT ACCOUNTED IN SRG BOOK" });
		categoryItems.put("TDS", new String[] { "TDS(2020-2021) ACCOUNTED ON XX.XX.2021"});

		Workbook workbook = new XSSFWorkbook();

		// hidden sheet for list values
		Sheet sheet = workbook.createSheet("ListSheet");
		CellStyle cellStyle = workbook.createCellStyle();
		cellStyle.setWrapText(true);
		
		Row row;
		Name namedRange;
		String colLetter;
		String reference;

		int c = 0;
		// put the data in
		for (String key : categoryItems.keySet()) {
			int r = 0;
			row = sheet.getRow(r);
			if (row == null)
				row = sheet.createRow(r);
			r++;
			row.createCell(c).setCellValue(key);
			String[] items = categoryItems.get(key);
			for (String item : items) {
				row = sheet.getRow(r);
				if (row == null)
					row = sheet.createRow(r);
				r++;
				row.createCell(c).setCellValue(item);
			}
			// create names for the item list constraints, each named from the current key
			colLetter = CellReference.convertNumToColString(c);
			namedRange = workbook.createName();
			namedRange.setNameName(key);
			reference = "ListSheet!$" + colLetter + "$2:$" + colLetter + "$" + r;
			namedRange.setRefersToFormula(reference);
			c++;
		}

		// create name for Categories list constraint
		colLetter = CellReference.convertNumToColString((c - 1));
		namedRange = workbook.createName();
		namedRange.setNameName("Categories");
		reference = "ListSheet!$A$1:$" + colLetter + "$1";
		namedRange.setRefersToFormula(reference);

		// unselect that sheet because we will hide it later
		sheet.setSelected(false);

		// visible data sheet
		sheet = workbook.createSheet("Sheet1");

		sheet.createRow(0).createCell(0).setCellValue("Transaction");
		sheet.getRow(0).createCell(1).setCellValue("Status");

		sheet.setActiveCell(new CellAddress("A2"));

		sheet.setColumnWidth(0,  5000); // Column A width is 20 characters
		sheet.setColumnWidth(1,  10000); // Column B width is 30 characters

		// data validations
		DataValidationHelper dvHelper = sheet.getDataValidationHelper();
		// data validation for categories in A2:
		DataValidationConstraint dvConstraint = dvHelper.createFormulaListConstraint("Categories");
		CellRangeAddressList addressList = new CellRangeAddressList(1, 1, 0, 0);
		DataValidation validation = dvHelper.createValidation(dvConstraint, addressList);
		sheet.addValidationData(validation);

		// data validation for items of the selected category in B2:

		dvConstraint = dvHelper.createFormulaListConstraint("INDIRECT($A$2)");
		System.out.println("dvConstraint" + dvConstraint);
		System.out.println("INDIRECT($A$2)");
		System.out.println();
		addressList = new CellRangeAddressList(1, 1, 1, 1);
		validation = dvHelper.createValidation(dvConstraint, addressList);
		sheet.addValidationData(validation);

		// hide the ListSheet
		workbook.setSheetHidden(0, true);
		// set Sheet1 active
		workbook.setActiveSheet(1);

		FileOutputStream out = new FileOutputStream("D:\\java\\PROJECTS\\FilterData.xlsx");
		workbook.write(out);
		System.out.println("Now the file is ready to open...");
		File file = new File("D:\\java\\PROJECTS\\FilterData.xlsx");
        try {
            Desktop.getDesktop().open(file);
        } catch (IOException e) {
            e.printStackTrace();
        }
		workbook.close();
		out.close();
		
	}
} // worked
