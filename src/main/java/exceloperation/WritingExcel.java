package exceloperation;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Map;
import java.util.TreeMap;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

//workbook-->sheet-->rows-->cells
public class WritingExcel {

	public static void main(String[] args) throws IOException {
		XSSFWorkbook workbook = new XSSFWorkbook();
		XSSFSheet sheet = workbook.createSheet("sheet");
		//Two dimensional object array
		Object student[][] = { { "#", "student_name", "mark1", "mark2", "mark3" }, 
				{ "1", "Anand", "64", "72", "43" },
				{ "2", "kavin", "76", "71", "55" }, 
				{ "3", "kavin", "50", "65", "58" },
				{ "4", "kavin", "77", "75", "69" }, 
				};

				int rows = student.length;
				int cols = student[0].length;
				System.out.println(rows);
				System.out.println(cols);

				for (int r = 0; r < rows; r++) {
					XSSFRow row = sheet.createRow(r);//create row
					for (int c = 0; c < cols; c++) {
						XSSFCell cell = row.createCell(c); // create column
						Object value = student[r][c];
						if (value instanceof String)
							cell.setCellValue((String) value);
						if (value instanceof Integer)
							cell.setCellValue((Integer) value);
						if (value instanceof Boolean)
							cell.setCellValue((Boolean) value);

					}
				}
				String filepath = ".\\datafiles\\student info.xlsx";
				FileOutputStream fos = new FileOutputStream(filepath);
				workbook.write(fos);
				fos.close();
				System.out.println("Data written successfully");

			}

		}
