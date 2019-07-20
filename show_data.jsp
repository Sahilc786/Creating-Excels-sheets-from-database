<%@page import="java.io.FileInputStream"%>
<%@page import="org.apache.poi.xssf.usermodel.XSSFWorkbook"%>
<%@ page language="java" contentType="text/html; charset=ISO-8859-1"
    pageEncoding="ISO-8859-1"%>
     <%@ page import="java.sql.*" %>
      <%@ page import="java.io.FileOutputStream" %>
     <%@ page import="org.apache.poi.ss.usermodel.Cell" %>
      <%@ page import="org.apache.poi.ss.usermodel.Row" %>
       <%@ page import="org.apache.poi.xssf.usermodel.XSSFSheet" %>
        <%@ page import="org.apache.poi.xssf.usermodel.XSSFWorkbook" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=ISO-8859-1">
<title>Insert title here</title>
</head>
<body>
<%
	XSSFWorkbook workbook=new XSSFWorkbook();
String d_n=request.getParameter("db_name");
String url="jdbc:mysql://localhost:3306/"+d_n;
Class.forName("com.mysql.jdbc.Driver");
Connection co=DriverManager.getConnection(url,"root","password");
Statement st=co.createStatement();
String q_table="show tables";
String T_n=" ",table="";
ResultSet rs1=st.executeQuery(q_table);

while(rs1.next())
{
	
	T_n=rs1.getString(1);
	XSSFSheet sheet1 = workbook.createSheet(T_n);
	String query_fetch="select * from "+T_n;
	Statement st2=co.createStatement();
	ResultSet rs2=st2.executeQuery(query_fetch);
	ResultSetMetaData rsmd=rs2.getMetaData();
	int col=rsmd.getColumnCount();
	Row row1 = sheet1.createRow(0);
	for(int i=1;i<=col;i++)
	{
		
		String name=rsmd.getColumnName(i);
		String type=rsmd.getColumnTypeName(i);
		
		Cell cell1 = row1.createCell(i-1);
		cell1.setCellValue(name+"("+type+")");
		sheet1.autoSizeColumn(i-1);
		
	}
	Row row2;
	Cell cell;
	int i=1;
	while(rs2.next())
	{
		
		row2=sheet1.createRow(i);
		for(int k=1;k<=col;k++)
		{
	
	
		       			cell = row2.createCell(k - 1);
				cell.setCellValue(rs2.getString(k));
				sheet1.autoSizeColumn(k - 1);
				

			}
			i++;
		}
		String path = "F:\\excel_sheets\\" + d_n + ".xlsx";
		FileOutputStream fo = new FileOutputStream(path);

		workbook.write(fo);

		fo.close();
	}

	workbook.close();
	out.println("EXCEL sheet created succesfully ");
%>
<br><br><br>
<h3><a href="index.jsp"><i style="color: red;">click here to go back</i> </a></h3>
</body>
