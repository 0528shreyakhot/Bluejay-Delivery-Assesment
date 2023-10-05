package utils;
import java.io.IOException;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.Scanner;
import java.text.SimpleDateFormat;
import java.time.LocalDateTime;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.time.format.DateTimeFormatter;
import java.time.LocalTime;
import java.time.Duration;
//import exceloperations.Employee;


class Employee {
    String Position_ID, name, Timecard,dateIn, timeIn, dateOut, timeOut;
	

    Employee(String Position_ID, String name2, String timecard2, String dateIn, String timeIn, String dateOut,
            String timeOut) {
        this.Position_ID = Position_ID;
        this.name = name2;
        this.Timecard = timecard2;
        this.dateIn = dateIn;
        this.timeIn = timeIn;
        this.dateOut = dateOut;
        this.timeOut = timeOut;
    }
}
public class ExcelUtils {

	
		public static void main(String []args) throws IOException, ParseException
		{
			getRowCount();
			getCellData();
		}
		
		public static void getCellData() throws IOException, ParseException
		{
			Scanner sc = new Scanner(System.in);
			List<Employee> employees = new ArrayList<Employee>();
			String excelPath = "./data/Assignment_Timecard.xlsx";
			XSSFWorkbook workbook = new XSSFWorkbook(excelPath);
			XSSFSheet sheet = workbook.getSheet("Sheet1");
			int rowCount = sheet.getPhysicalNumberOfRows();
			int choice;
			DataFormatter formatter = new DataFormatter();
			
			DateTimeFormatter formater = DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm:ss");			
			int count=0;
			for(int i=2; i<= rowCount-3; i++)
			{
				String Position_ID = sheet.getRow(i).getCell(0).getStringCellValue();
				
				String name = sheet.getRow(i).getCell(7).getStringCellValue();
				Object temp = formatter.formatCellValue(sheet.getRow(i).getCell(4));
				String Timecard = String.valueOf(temp);

				temp = formatter.formatCellValue(sheet.getRow(i).getCell(2));
				String dateTimeIn = String.valueOf(temp);

                SimpleDateFormat inputFormat = new SimpleDateFormat("yyyy/MM/dd hh:mm a");
                Date date = inputFormat.parse(dateTimeIn);

                SimpleDateFormat dateFormat = new SimpleDateFormat("yyyy/MM/dd");
                SimpleDateFormat timeFormat = new SimpleDateFormat("hh:mm a");

                String dateStrIn = dateFormat.format(date);
                String timeStrIn = timeFormat.format(date);

            
                temp = formatter.formatCellValue(sheet.getRow(i).getCell(2));
                String dateTimeOut = String.valueOf(temp);
                 inputFormat = new SimpleDateFormat("yyyy/MM/dd hh:mm a");
                 date = inputFormat.parse(dateTimeOut);

                 dateFormat = new SimpleDateFormat("yyyy/MM/dd");
                 timeFormat = new SimpleDateFormat("hh:mm a");

                String dateStrOut = dateFormat.format(date);
                String timeStrOut = timeFormat.format(date);

             employees.add(new Employee(Position_ID, name, Timecard, dateStrIn, timeStrIn, dateStrOut, timeStrOut));

			}
			
			
			
			while(true)
			{
				System.out.println("Choose the Option : ");
				System.out.println("\t1) Who has worked for 7 consecutive days\n"
						+ "	2) Who have less than 10 hours of time between shifts but greater than 1 hour\n\t"
						+ "3) Who has worked for more than 14 hours in a single shift");
				
				choice = sc.nextInt();
				if(choice == 1)
				{
					consecutivedays(employees);
				}
				else if(choice == 2)
				{
					differencebetsheefts(employees);
				}
				else if(choice == 3)
				{
					forteenhours(employees);
				}
				
				else
				{
					System.out.println("wrong input");
				}
			}
			
			
		}
		
		
		private static void exit(int i) {
			// TODO Auto-generated method stub
			
		}

		public static void consecutivedays(List<Employee> employees)
		{
			System.out.println("Workers who has worked for 7 consecutive days are : ");
			System.out.println("-------------------------------------------------------------------------------");
			
			
			int size =employees.size();
			int c=1;
			for(int i = 1; i<size-9;i++)
			{
				c=1;
				for(int j = i+2; j<size-9;j++)
				{

					if(employees.get(i).Position_ID == employees.get(j).Position_ID)
					{
						
						c++;
						j++;
						if(c == 7)
						{
							System.out.println("Position_ID = "+ employees.get(i).Position_ID);
							System.out.println("Name = "+ employees.get(i).name);
							System.out.println("-------------------------------------------------------------------------------");
						}
					}
					else
					{
						break;
					}
					
				}
				i = i+c*2;
			}
		}
		
		public static void differencebetsheefts(List<Employee> employees) throws ParseException 
		{
			System.out.println("Workers  who have less than 10 hours of time between shifts but greater than 1 hour:  ");
			System.out.println("-------------------------------------------------------------------------------");
			int c=2;
			int size =employees.size();
			for(int i = 2; i<size-11;i++)
			
			{
	

					String time1Str = employees.get(i).timeIn;
		            String time2Str = employees.get(i+1).timeOut;
		            
		            SimpleDateFormat sdf = new SimpleDateFormat("hh:mm a");
		            
		            Date time1 = sdf.parse(time1Str);
		            Date time2 = sdf.parse(time2Str);
		            
		            long timeDifferenceMillis = time2.getTime() - time1.getTime();
		            long hours = timeDifferenceMillis / (60 * 60 * 1000);
		            if(hours <= 10 && hours >=1)
						 {
						        	System.out.println("Position_ID = "+ employees.get(i).Position_ID);
									System.out.println("Name = "+ employees.get(i).name);
									System.out.println("-------------------------------------------------------------------------------");
						 }

				}
		}
		
		
		public static void forteenhours(List<Employee> employees) throws ParseException
		{
			System.out.println("Workers Who has worked for more than 14 hours in a single shift : ");
			System.out.println("-------------------------------------------------------------------------------");
			int c=2;
			int size =employees.size();
			for(int i = 2; i<size-11;i++)
			
			{
	

					String time1Str = employees.get(i).timeIn;
		            String time2Str = employees.get(i+1).timeOut;
		            
		            SimpleDateFormat sdf = new SimpleDateFormat("hh:mm a");
		            
		            Date time1 = sdf.parse(time1Str);
		            Date time2 = sdf.parse(time2Str);
		            
		            long timeDifferenceMillis = time2.getTime() - time1.getTime();
		            long hours = timeDifferenceMillis / (60 * 60 * 1000);
		            if(hours >= 14)
						 {
						        	System.out.println("Position_ID = "+ employees.get(i).Position_ID);
									System.out.println("Name = "+ employees.get(i).name);
									System.out.println("-------------------------------------------------------------------------------");
						 }

			}

			}
			
		
		public static void getRowCount()
		{
			try{
				String excelPath = "./data/Assignment_Timecard.xlsx";
				XSSFWorkbook workbook = new XSSFWorkbook(excelPath);
				XSSFSheet sheet = workbook.getSheet("Sheet1");
				int rowCount = sheet.getPhysicalNumberOfRows();

				
			}catch(Exception exp)
			{
				System.out.println(exp.getCause());
				System.out.println(exp.getMessage());
				exp.printStackTrace();
			}
		}
}
