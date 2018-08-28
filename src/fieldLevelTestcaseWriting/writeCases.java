package fieldLevelTestcaseWriting;

import fieldLevelTestcaseWriting.getData;
import resources.steps;
import resources.rowNumberHolder;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashMap;


public class writeCases {
	
	private static final String FILE_NAME = "K:\\Project\\fieldLevelTestcaseWriting\\testFolder\\IDD_ISP_Testcases.xlsx";
	private static final getData data = new getData();
	private static final steps Steps = new steps();
	private static XSSFWorkbook workbook = new XSSFWorkbook();
	private static XSSFSheet sheet = workbook.createSheet("DOC");
	
	public static void main(String[] args) {
		

		data.setDataSheet();
		int rowCount = 1;
		Row row = sheet.createRow(rowCount);
		for(int dataCount = 1; dataCount <= data.getNumberOfCases(); dataCount++) {
			data.setCurrentRow(dataCount);
			HashMap<Integer, String[][]> EverySteps = new HashMap<Integer, String[][]>();
			Integer totalCasesCount = 1;
			data.isMaxLengthFixed();
			Steps.setTcSpecificValues(data.getMaxLengthOfField(), data.getFieldName(), data.getFieldType(), data.getDefaultValue(), data.getFieldValue(), data.getEnableDisable());
			EverySteps.put(totalCasesCount++, Steps.getTestCaseSteps("steps"));
			EverySteps.put(totalCasesCount++, Steps.getTestCaseSteps(data.getFieldType()));
			if (data.isMaxLengthFixed()) {
				EverySteps.put(totalCasesCount++, Steps.getTestCaseSteps("length"));
			}
			if (data.checkIsItMandatory()) {
				EverySteps.put(totalCasesCount++, Steps.getTestCaseSteps("mandatory"));
			}
			int StepCount = 1;
			row.createCell(rowNumberHolder.TC_ID_ROW_NUMBER).setCellValue(data.getTcId());
			row.createCell(rowNumberHolder.USE_CASE_ROW_NUMBER).setCellValue(data.getUsecase());
			row.createCell(rowNumberHolder.FR_ROW_NUMBER).setCellValue(data.getFr());
			row.createCell(rowNumberHolder.BR_ROW_NUMBER).setCellValue(data.getBr());
			row.createCell(rowNumberHolder.TC_TYPE_ROW_NUMBER).setCellValue(data.getTcType());
			row.createCell(rowNumberHolder.DESCRIPTION_ROW_NUMBER).setCellValue(data.getDescription(data.getFieldName()));
			for(int diffCasesCount = 1; diffCasesCount < totalCasesCount; diffCasesCount++) {
				String[][] Step = EverySteps.get(diffCasesCount);
				int totalRows = 0;
				for(; totalRows < Step[0].length; totalRows++) {
					row.createCell(rowNumberHolder.STEP_COUNT_ROW_NUMBER).setCellValue("Step "+String.valueOf(StepCount));
					row.createCell(rowNumberHolder.STEP_ROW_NUMBER).setCellValue(Step[0][totalRows]);
					row.createCell(rowNumberHolder.EXPECT_RESULT_ROW_NUMBER).setCellValue(Step[1][totalRows]);
					StepCount++;
					rowCount++;
					row = sheet.createRow(rowCount);
				}
			}
			sheet = mergeCells(rowCount,StepCount,sheet);
		}
		
		try {
            FileOutputStream outputStream = new FileOutputStream(FILE_NAME);
            workbook.write(outputStream);
            workbook.close();
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }

        System.out.println("Done");
	}
	
	public static XSSFSheet mergeCells(int rowCount, int StepCount, XSSFSheet sheet) {
		sheet.addMergedRegion(new CellRangeAddress(rowCount-StepCount+1,rowCount-1,rowNumberHolder.TC_ID_ROW_NUMBER,rowNumberHolder.TC_ID_ROW_NUMBER));
		sheet.addMergedRegion(new CellRangeAddress(rowCount-StepCount+1,rowCount-1,rowNumberHolder.USE_CASE_ROW_NUMBER,rowNumberHolder.USE_CASE_ROW_NUMBER));
		sheet.addMergedRegion(new CellRangeAddress(rowCount-StepCount+1,rowCount-1,rowNumberHolder.FR_ROW_NUMBER,rowNumberHolder.FR_ROW_NUMBER));
		sheet.addMergedRegion(new CellRangeAddress(rowCount-StepCount+1,rowCount-1,rowNumberHolder.BR_ROW_NUMBER,rowNumberHolder.BR_ROW_NUMBER));
		sheet.addMergedRegion(new CellRangeAddress(rowCount-StepCount+1,rowCount-1,rowNumberHolder.TC_TYPE_ROW_NUMBER,rowNumberHolder.TC_TYPE_ROW_NUMBER));
		sheet.addMergedRegion(new CellRangeAddress(rowCount-StepCount+1,rowCount-1,rowNumberHolder.DESCRIPTION_ROW_NUMBER,rowNumberHolder.DESCRIPTION_ROW_NUMBER));
		return sheet;
	}
	
}
