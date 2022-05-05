package read.write.operation;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Read_Write_Operation {

	public static  void main(String args[]) throws IOException {

		File file =    new File("D:\\test123xlsx.xlsx");

		FileInputStream inputStream = new FileInputStream(file);

		try (XSSFWorkbook wb = new XSSFWorkbook(inputStream)) {
			XSSFSheet sheet = wb.getSheet("test");

			int rowCount=sheet.getLastRowNum()-sheet.getFirstRowNum();
			
			for(int i=0;i<=rowCount;i++){

				//get cell count in a row
				int cellcount=sheet.getRow(i).getLastCellNum();

				//iterate over each cell to print its value
				System.out.println("Row "+ i+" data is :");

				for(int j=0;j<cellcount;j++){
					System.out.print(sheet.getRow(i).getCell(j).getStringCellValue() +",");
					//System.out.println(sheet.getRow(i).getCell(j).getNumericCellValue()+",");
				}
				System.out.println();
			}
		}
	}
}
