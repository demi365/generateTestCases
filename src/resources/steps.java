package resources;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.HashMap;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class steps {
	
	private static HashMap<String,String[][]> typeWiseStep = new HashMap<String,String[][]>();
	private Integer TcSpecificLength = 0;
	private String TcSpecificFieldName = "";
	private String TcSpecificFieldType = "";
	private String TcSpecificDefaultValue = "";
	private String TcSpecificFieldValue = "";
	private String TcSpecificEnableDisable = "";
	
	private static final String STEP_FILE_NAME = "K:\\Project\\fieldLevelTestcaseWriting\\resources\\test_steps.xlsx";
	private static final String GEN_SHEET_NAME = "general steps";
	private static final String STEP_SHEET_NAME = "field steps";
	private static FileInputStream dataFile = null;
	private static Workbook dataWorkbook = null;
	private static Sheet dataSheet = null;
	private static Row currentRow = null;
	
	
	public steps() {
		try {
			dataFile = new FileInputStream(new File(STEP_FILE_NAME));
			dataWorkbook = new XSSFWorkbook(dataFile);
			
			/* Setting the first initial steps for the test cases*/
			
			int dataRow=1;
			dataSheet = dataWorkbook.getSheet(GEN_SHEET_NAME);
			currentRow = dataSheet.getRow(dataRow);
			
			int lenOfCurrentStep = lenOfCurrentStep(currentRow, dataRow);
			String currentStep[][] = new String[2][lenOfCurrentStep];
			
			for(int len=0; len< lenOfCurrentStep; len++) {
				currentStep[0][len] = currentRow.getCell(1).getStringCellValue();
				currentStep[1][len] = currentRow.getCell(2).getStringCellValue();
				
				dataRow++;
				currentRow = dataSheet.getRow(dataRow);
			}
			typeWiseStep.put("steps", currentStep);
			
			
			/* Setting the next field level steps for the test cases*/
			
			dataRow=1;
			dataSheet = dataWorkbook.getSheet(STEP_SHEET_NAME);
			currentRow = dataSheet.getRow(dataRow);
			while(!currentRow.getCell(0).getStringCellValue().equalsIgnoreCase("_end_")) {
				String currentFieldType = currentRow.getCell(0).getStringCellValue().toLowerCase();
				lenOfCurrentStep = lenOfCurrentStep(currentRow, dataRow);
				currentStep = new String[2][lenOfCurrentStep];
				System.out.println(lenOfCurrentStep);

				for(int len=0; len< lenOfCurrentStep; len++) {
					currentStep[0][len] = currentRow.getCell(1).getStringCellValue();
					currentStep[1][len] = currentRow.getCell(2).getStringCellValue();
					
					dataRow++;
					currentRow = dataSheet.getRow(dataRow);
				}
				System.out.println(currentFieldType);
				
				for(int i = 0; i<2; i++) {
					for(int j =0; j<lenOfCurrentStep; j++) {
						System.out.println(currentStep[i][j]);
					}
				}
				typeWiseStep.put(currentFieldType, currentStep);
			}
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
	}
	
	@SuppressWarnings("deprecation")
	public static int lenOfCurrentStep(Row currentRow, int rowNum) {
		int lenOfCurrentStepValue = 1;
		currentRow = dataSheet.getRow(rowNum+lenOfCurrentStepValue);
		while((currentRow.getCell(0) == null || currentRow.getCell(0).getCellType() == Cell.CELL_TYPE_BLANK)) {
			lenOfCurrentStepValue++;
			currentRow = dataSheet.getRow(rowNum+lenOfCurrentStepValue);
		}
		return lenOfCurrentStepValue;
	}
	
	public String[][] getTestCaseSteps(String type_name){
		return this.removePlaceHolders(typeWiseStep.get(type_name.toLowerCase()));
	}
	
	public void setTcSpecificValues(int maxlength, String FieldName, String FieldType, String DefaultValue, String FieldValue, String EnableDisable) {
		TcSpecificLength = maxlength;
		TcSpecificFieldName = FieldName;
		TcSpecificFieldType = FieldType;
		TcSpecificDefaultValue = DefaultValue;
		TcSpecificFieldValue = FieldValue;
		TcSpecificEnableDisable = EnableDisable;
		System.out.println(TcSpecificFieldName+TcSpecificFieldType);
	}
	
	public String[][] removePlaceHolders(String[][] TcWithPlaceHolders){
		String[][] TcAfterRemovingPlaceHolders = new String[2][TcWithPlaceHolders[0].length];
		for(int leng = 0; leng < TcWithPlaceHolders[0].length; leng++) {
			TcAfterRemovingPlaceHolders[0][leng] = TcWithPlaceHolders[0][leng].replaceAll("<<field>>", TcSpecificFieldName).replaceAll("<<Field_Type>>", TcSpecificFieldType).replaceAll("<<char>>", String.valueOf(TcSpecificLength)).replaceAll("<<value>>", TcSpecificFieldValue).replaceAll("<<default value>>", TcSpecificDefaultValue);
			TcAfterRemovingPlaceHolders[1][leng] = TcWithPlaceHolders[1][leng].replaceAll("<<field>>", TcSpecificFieldName).replaceAll("<<Field_Type>>", TcSpecificFieldType).replaceAll("<<char>>", String.valueOf(TcSpecificLength)).replaceAll("<<value>>", TcSpecificFieldValue).replaceAll("<<default value>>", TcSpecificDefaultValue);
		}
		return TcAfterRemovingPlaceHolders;
	}
	
}
