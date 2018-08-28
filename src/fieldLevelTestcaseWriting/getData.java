package fieldLevelTestcaseWriting;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class getData {
	
	private static final String UseCase = "N/A";
	private static final String FR = "N/A";
	private static final String BR = "N/A";
	private static final String TC_type = "Field Validation";
	private static final String Description = "Verify the functional validation of the <<field>> field.";
	private static final String idd_module = "TC_IDD_DOC_";
	private static final String UI_SPEC_FILE_NAME = "K:\\Project\\fieldLevelTestcaseWriting\\resources\\Test_MockUp_UI_Field_level_specs.xlsx";
	private static final String UI_SPEC_SHEET_NAME = "S-1";
	private static final Integer FieldNameColumn = 1;
	private static final Integer FieldTypeColumn = 2;
	private static final Integer FieldMandatoryColumn = 3;
	private static final Integer FieldMaxLengthColumn = 4;
	private static final Integer FieldValueColumn = 5;
	private static final Integer DefaultValueColumn = 6;
	private static final Integer EnableDisableColumn = 7;
	
	private static final Integer FieldStartRow = 3;
	private static final Integer FieldEndRow = 11;
	
	private static Integer maxLengthOfField = 0;
	
	private static Integer counter = 0;
	private static FileInputStream dataFile = null;
	private static Workbook dataWorkbook = null;
	private static Sheet dataSheet = null;
	private static Row currentRow = null;
	
	public void setDataSheet() {
		try {
			dataFile = new FileInputStream(new File(UI_SPEC_FILE_NAME));
			dataWorkbook = new XSSFWorkbook(dataFile);
			dataSheet = dataWorkbook.getSheet(UI_SPEC_SHEET_NAME);
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
	}
	
	public void setCurrentRow(int increment) {
		System.out.println("Setting current row : "+String.valueOf(FieldStartRow+increment-1));
		currentRow = dataSheet.getRow(FieldStartRow+increment-1);
	}
	
	@SuppressWarnings("deprecation")
	public String getFieldValue() {
		if (currentRow.getCell(FieldValueColumn) != null && currentRow.getCell(FieldValueColumn).getCellType() != Cell.CELL_TYPE_BLANK) {
			String FieldValue = currentRow.getCell(FieldValueColumn).getStringCellValue();
			System.out.println("The field value is : "+FieldValue);
			return FieldValue;
		}
		return "";
	}

	@SuppressWarnings("deprecation")
	public String getDefaultValue() {
		if (currentRow.getCell(DefaultValueColumn) != null && currentRow.getCell(DefaultValueColumn).getCellType() != Cell.CELL_TYPE_BLANK) {
			String DefaultValue = currentRow.getCell(DefaultValueColumn).getStringCellValue();
			System.out.println("The Default value is : "+DefaultValue);
			return DefaultValue;
		}
		return "";
	}

	@SuppressWarnings("deprecation")
	public String getEnableDisable() {
		if (currentRow.getCell(EnableDisableColumn) != null && currentRow.getCell(EnableDisableColumn).getCellType() != Cell.CELL_TYPE_BLANK) {
			String EnableDisable = currentRow.getCell(EnableDisableColumn).getStringCellValue();
			System.out.println("The Editable status is : "+EnableDisable);
			return EnableDisable;
		}
		return "";
	}

	@SuppressWarnings("deprecation")
	public String getFieldName() {
		if (currentRow.getCell(FieldNameColumn) != null && currentRow.getCell(FieldNameColumn).getCellType() != Cell.CELL_TYPE_BLANK) {
			String FieldName = currentRow.getCell(FieldNameColumn).getStringCellValue();
			System.out.println("The field name is : "+FieldName);
			return FieldName;
		}
		return "";
	}
	
	public String getFieldType() {
		String Field_Type = currentRow.getCell(FieldTypeColumn).getStringCellValue().toLowerCase();
		System.out.println("The field type is : "+Field_Type);
		return Field_Type;
	}
	
	public boolean checkIsItMandatory() {
		if (currentRow.getCell(FieldMandatoryColumn).getStringCellValue().toLowerCase().contains("y"))
			return true;
		return false;
	}
	
	@SuppressWarnings("deprecation")
	public boolean isMaxLengthFixed() {
		if (currentRow.getCell(FieldMaxLengthColumn) != null && currentRow.getCell(FieldMaxLengthColumn).getCellType() == Cell.CELL_TYPE_NUMERIC) {
			maxLengthOfField = (int) currentRow.getCell(FieldMaxLengthColumn).getNumericCellValue();
			return true;
		}
		return false;
	}
	
	public Row getCurrentRow() {
		return currentRow;
	}
	
	public int getNumberOfCases() {
		System.out.println("Total number of cases :"+String.valueOf(FieldEndRow-FieldStartRow));
		return FieldEndRow-FieldStartRow;
	}
	
	public String getTcId() {
		counter++;
		String TcId  = idd_module + counter.toString();
		return TcId;
	}

	public String getFr() {
		return FR;
	}

	public String getBr() {
		return BR;
	}

	public String getTcType() {
		return TC_type;
	}

	public String getUsecase() {
		return UseCase;
	}

	public String getDescription(String FieldName) {
		return Description.replaceAll("<<field>>", FieldName);
	}

	public Integer getMaxLengthOfField() {
		return maxLengthOfField;
	}
	
}
