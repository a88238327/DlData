package service;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.OutputStream;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;



public class memberData{

	public static void main(String[] args) throws Exception {
		HSSFWorkbook wb=new HSSFWorkbook();	
		HSSFSheet sheet=wb.createSheet();
		HSSFRow row= sheet.createRow(0);
		HSSFCell cell=row.createCell(0);
		cell.setCellValue("halo");
		ResultSet rs=null;
		String[] tiltl=new String[] {"创建时间","微信名","手机号","邀请码"};
		String filename="会员信息";
		Connection conn=null;
		Statement stmt=null;
		PreparedStatement pstmt	= null ;
		String driver="com.microsoft.sqlserver.jdbc.SQLServerDriver";//驱动类
		String username="sa";//数据库用户名
		String password="88238327";//数据库密码
		String sql="select createtime,nickname,USER_FORM_INFO_FLAG_MOBILE,邀请码 from customer" ;//查询语句
		String url="jdbc:sqlserver://localhost:1433;DatabaseName=chejiacloud";//连接数据库的地址
		try{
			Class.forName(driver);//加载驱动器类
			conn=DriverManager.getConnection(url,username,password);//建立连接
			//建立处理的SQL语句
			pstmt = conn.prepareStatement(sql) ;
			rs = pstmt.executeQuery() ;//形成结果集
			
			if(memberdata(rs, tiltl, filename))
			{
				System.out.println("success");
			}
			rs.close();//关闭结果集
			pstmt.close();//关闭SQL语句集
			conn.close();//关闭连接
		}catch (Exception e) {
			// TODO: handle exception
			System.out.print(e);
		}
	}
	public static boolean memberdata(ResultSet rs,String[] tiltl,String filename) throws Exception {
		HSSFWorkbook wb=new HSSFWorkbook();
		HSSFSheet sheet=wb.createSheet();
		HSSFRow row=sheet.createRow(0);
		HSSFCell cell=row.createCell(0);
		for(int i=0;i<tiltl.length;i++)
		{
			cell=row.createCell(i);
			cell.setCellValue(tiltl[i]);
		}
		for(int i=1;rs.next();i++)
		{
			row=sheet.createRow(i);
			for(int j=0;j<tiltl.length;j++)
			{
				cell=row.createCell(j);
				cell.setCellValue(rs.getString(j+1));
				System.out.println(rs.getString(j+1));
			}
		}
		File file=new File("c:\\会员卡.xls");
		file.createNewFile();
				
		OutputStream flieOut=new FileOutputStream("c:\\会员卡.xls");
		wb.write(flieOut);
		wb.close();
		flieOut.close();
		return true;
	}
}
