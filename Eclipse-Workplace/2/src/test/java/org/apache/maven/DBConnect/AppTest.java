package org.apache.maven.DBConnect;

import java.io.File;

import java.io.IOException;
import java.sql.Connection;

import java.sql.DriverManager;

import java.sql.SQLException;
import java.sql.Statement;
import java.util.concurrent.ExecutionException;

import org.testng.annotations.Test;

import it.firegloves.mempoi.MemPOI;
import it.firegloves.mempoi.builder.MempoiBuilder;
import it.firegloves.mempoi.builder.MempoiSheetBuilder;
import it.firegloves.mempoi.domain.MempoiSheet;

public class AppTest {
	Statement st = null;
	Connection conn = null;

	@Test
	public void DB() throws IOException, SQLException, InterruptedException, ExecutionException {

		ExcelUtils1.SetExcelFile("E://Execute.xlsx", "Sheet1");
		System.out.println(ExcelUtils1.getTotalRows());
		for (int j = 1; j <= ExcelUtils1.getTotalRows(); j++) {

			String ProjectName = ExcelUtils1.getCellData(j, 0);
			String DBName = ExcelUtils1.getCellData(j, 1);

			try {
				DriverManager.registerDriver(new com.microsoft.sqlserver.jdbc.SQLServerDriver());
				// Connection con=
				// DriverManager.getConnection("jdbc:mysql://OUSPXD19:1433/naa_mobile_systems","td","AHWg69amjWrA5pV");

				String dbURL = "jdbc:sqlserver://10.233.18.31:1433;user=td;password=tdtdtd;database=" + DBName + "";
				conn = DriverManager.getConnection(dbURL);

				if (conn != null) {
					System.out.println("Connected");
				}
			} catch (Exception e) {
				System.out.println("issue");
			}

			ExcelUtils.SetExcelFile("E://Execute.xlsx", "Sheet2");

			String BugAttachment = ExcelUtils.getCellData(1, 1);
			String DesignStepAtt = ExcelUtils.getCellData(2, 1);

			String ReqAtt = ExcelUtils.getCellData(3, 1);
			String RunStepAtt = ExcelUtils.getCellData(4, 1);
			String TestAtt = ExcelUtils.getCellData(5, 1);
			String Requirement = ExcelUtils.getCellData(6, 1);
			String TestPlan = ExcelUtils.getCellData(7, 1);
			String TestResult = ExcelUtils.getCellData(8, 1);
			String TestExecution = ExcelUtils.getCellData(9, 1);
			String TestSet_NOReq = ExcelUtils.getCellData(10, 1);
			String TestSet_withReq = ExcelUtils.getCellData(11, 1);
			String Test_Folder = ExcelUtils.getCellData(12, 1);
			String RunAttachment = ExcelUtils.getCellData(13, 1);
			String TestSetAttachment = ExcelUtils.getCellData(14, 1);
			String TestSetAttachment1 = ExcelUtils.getCellData(15, 1);

			File fileDest = new File("E:\\" + ProjectName + "\\" + ProjectName + ".xlsx");

			MempoiSheet Bugsheet = MempoiSheetBuilder.aMempoiSheet().withSheetName("BugAttachment")
					.withPrepStmt(conn.prepareStatement(BugAttachment)).build();

			MempoiSheet DesignStepSheet = MempoiSheetBuilder.aMempoiSheet().withSheetName("DesignStepAtt")
					.withPrepStmt(conn.prepareStatement(DesignStepAtt)).build();

			MempoiSheet ReqAttSheet = MempoiSheetBuilder.aMempoiSheet().withSheetName("ReqAtt")
					.withPrepStmt(conn.prepareStatement(ReqAtt)).build();

			MempoiSheet RunStepSheet = MempoiSheetBuilder.aMempoiSheet().withSheetName("RunStepAtt")
					.withPrepStmt(conn.prepareStatement(RunStepAtt)).build();

			MempoiSheet TestAttSheet = MempoiSheetBuilder.aMempoiSheet().withSheetName("TestAtt")
					.withPrepStmt(conn.prepareStatement(TestAtt)).build();

			MempoiSheet ReqSheet = MempoiSheetBuilder.aMempoiSheet().withSheetName("Requirement")
					.withPrepStmt(conn.prepareStatement(Requirement)).build();

			MempoiSheet TestSheet = MempoiSheetBuilder.aMempoiSheet().withSheetName("TestPlan")
					.withPrepStmt(conn.prepareStatement(TestPlan)).build();

			MempoiSheet TestResultSheet = MempoiSheetBuilder.aMempoiSheet().withSheetName("TestResult")
					.withPrepStmt(conn.prepareStatement(TestResult)).build();

			MempoiSheet TestExecutionSheet = MempoiSheetBuilder.aMempoiSheet().withSheetName("TestExecution")
					.withPrepStmt(conn.prepareStatement(TestExecution)).build();

			MempoiSheet TestSet_NOReqSheet = MempoiSheetBuilder.aMempoiSheet().withSheetName("TestSet_NOReq")
					.withPrepStmt(conn.prepareStatement(TestSet_NOReq)).build();

			MempoiSheet TestSet_withReqSheet = MempoiSheetBuilder.aMempoiSheet().withSheetName("TestSet_withReq")
					.withPrepStmt(conn.prepareStatement(TestSet_withReq)).build();

			MempoiSheet Test_FolderSheet = MempoiSheetBuilder.aMempoiSheet().withSheetName("Test_Folder")
					.withPrepStmt(conn.prepareStatement(Test_Folder)).build();

			MempoiSheet RunAttachmentSheet = MempoiSheetBuilder.aMempoiSheet().withSheetName("RunAttachment")
					.withPrepStmt(conn.prepareStatement(RunAttachment)).build();

			MempoiSheet TestSetAttachmentSheet = MempoiSheetBuilder.aMempoiSheet().withSheetName("TestSetAttachment")
					.withPrepStmt(conn.prepareStatement(TestSetAttachment)).build();

			MempoiSheet TestSetAttachment1Sheet = MempoiSheetBuilder.aMempoiSheet().withSheetName("TestSetAttachment1")
					.withPrepStmt(conn.prepareStatement(TestSetAttachment1)).build();

			MemPOI memPOI = MempoiBuilder.aMemPOI()

					.withFile(fileDest).withAdjustColumnWidth(true).addMempoiSheet(Bugsheet)
					.addMempoiSheet(DesignStepSheet)

					.addMempoiSheet(ReqAttSheet).addMempoiSheet(RunStepSheet).addMempoiSheet(TestAttSheet)
					.addMempoiSheet(ReqSheet).addMempoiSheet(TestSheet).addMempoiSheet(TestResultSheet)
					.addMempoiSheet(TestExecutionSheet).addMempoiSheet(TestSet_NOReqSheet)
					.addMempoiSheet(TestSet_withReqSheet).addMempoiSheet(Test_FolderSheet)
					.addMempoiSheet(RunAttachmentSheet).addMempoiSheet(TestSetAttachmentSheet)
					.addMempoiSheet(TestSetAttachment1Sheet)

					.build();

			// exports to file and gets the generated report absolute filename
			String absFilename = memPOI.prepareMempoiReportToFile().get();

			conn.close();

		}
	}
}

/*
 * HSSFRow rowhead = sheet.createRow((short) 0);
 * 
 * for (int k = 0; k <= numberOfColumns; k++) { String columnName =
 * metaData.getColumnName(k); Cell headerCell = rowhead.createCell(k);
 * headerCell.setCellValue(columnName); }
 * 
 * rowhead.createCell((short) 0).setCellValue("CellHeadName1");
 * rowhead.createCell((short) 1).setCellValue("CellHeadName2");
 * rowhead.createCell((short) 2).setCellValue("CellHeadName3"); int j = 1; while
 * (result.next()){ HSSFRow row = sheet.createRow((short) j);
 * 
 * for (int l = 0; l <= numberOfColumns; l++) { String columnName =
 * metaData.getColumnName(i); Cell headerCell = row.createCell(i); Object
 * valueObject = result.getObject(l); if (valueObject instanceof Integer) {
 * headerCell.setCellValue((Integer) valueObject); }else
 * headerCell.setCellValue((String) valueObject);
 * 
 * headerCell.setCellValue(columnName); } row.createCell((short)
 * 0).setCellValue(Integer.toString(result.getInt("column1")));
 * row.createCell((short) 1).setCellValue(result.getString("column2"));
 * row.createCell((short) 2).setCellValue(result.getString("column3")); j++; }
 * String yemi = "E:\\test.xls"; FileOutputStream fileOut = new
 * FileOutputStream(yemi); workbook.write(fileOut); fileOut.close(); } catch
 * (SQLException e1) { e1.printStackTrace(); } catch (FileNotFoundException e1)
 * { e1.printStackTrace(); } catch (IOException e1) { e1.printStackTrace(); }
 * 
 * } }
 */

/*
 * private void writeDataLines(ResultSet result, XSSFWorkbook workbook,
 * XSSFSheet sheet) throws SQLException { // write header line containing column
 * names ResultSetMetaData metaData = result.getMetaData(); int numberOfColumns
 * = metaData.getColumnCount();
 * 
 * Row headerRow = sheet.createRow(1);
 * 
 * // exclude the first column which is the ID field for (int i = 0; i <=
 * numberOfColumns; i++) { String columnName = metaData.getColumnName(i); Cell
 * headerCell = headerRow.createCell(i); headerCell.setCellValue(columnName); }
 * 
 * }
 * 
 * private void writeHeaderLine(ResultSet result, XSSFSheet sheet) throws
 * SQLException { ResultSetMetaData metaData = result.getMetaData(); int
 * numberOfColumns = metaData.getColumnCount();
 * 
 * int rowCount = 1;
 * 
 * while (result.next()) { Row row = sheet.createRow(rowCount++);
 * 
 * for (int i = 2; i <= numberOfColumns; i++) { Object valueObject =
 * result.getObject(i);
 * 
 * Cell cell = row.createCell(i - 2);
 * 
 * if (valueObject instanceof Integer) { cell.setCellValue((Integer)
 * valueObject); }else cell.setCellValue((String) valueObject);
 * 
 * }
 * 
 * }
 */
