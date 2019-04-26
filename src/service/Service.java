package service;

import java.io.File;
import java.io.FileOutputStream;
import java.io.OutputStream;
import java.net.URLEncoder;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.Statement;

import javax.servlet.http.HttpServletResponse;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import net.sf.json.JSONArray;
import net.sf.json.JSONObject;

public class Service {
	public static void downloadData(String startime,String endtime,HttpServletResponse response) throws Exception {
		HSSFWorkbook wb=new HSSFWorkbook();	
		HSSFSheet sheet=wb.createSheet();
		HSSFRow row= sheet.createRow(0);
		HSSFCell cell=row.createCell(0);
		ResultSet rs=null;
		String[] tiltl=new String[] {"序号","创建时间","微信名","手机号","二维码归属","上级账号","公众号"};
		String filename="会员激活关注信息信息"+startime+"-"+endtime;
		Connection conn=null;
		Statement stmt=null;
		PreparedStatement pstmt	= null ;
		String driver="com.microsoft.sqlserver.jdbc.SQLServerDriver";//驱动类
		String username="sa";//数据库用户名
		String password="88238327";//数据库密码
		String sql="select customer.createtime,customer.nickname,customer.USER_FORM_INFO_FLAG_MOBILE,customer.OuterStr,manager.leader,customer.openid from customer left join manager on customer.OuterStr=manager.phone_number where customer.createtime between '"+startime+" 00:00:00' and '"+endtime+" 24:60:00' and USER_FORM_INFO_FLAG_MOBILE is not null";//查询语句
		String url="jdbc:sqlserver://localhost:1433;DatabaseName=chejiacloud";//连接数据库的地址
		try{
			Class.forName(driver);//加载驱动器类
			conn=DriverManager.getConnection(url,username,password);//建立连接
			//建立处理的SQL语句		
			pstmt = conn.prepareStatement(sql) ;		
			rs = pstmt.executeQuery() ;//形成结果集	
			if(memberdata(rs, tiltl, filename,response))
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
	public static boolean memberdata(ResultSet rs,String[] tiltl,String filename,HttpServletResponse response) throws Exception {
		HSSFWorkbook wb=new HSSFWorkbook();
		HSSFSheet sheet=wb.createSheet();
		HSSFRow row=sheet.createRow(0);
		HSSFCell cell=row.createCell(0);
		String url="https://api.weixin.qq.com/cgi-bin/user/get?access_token=ACCESS_TOKEN";
		String GET_TOKEN_URL=Util.get("http://cloud.hnjtbf.com/wechat/gettoken");
		String resultString=Util.get(url.replace("ACCESS_TOKEN", GET_TOKEN_URL));
		JSONObject jsonObject=JSONObject.fromObject(resultString);
		String a=JSONObject.fromObject(jsonObject.getString("data")).getString("openid");
		JSONArray jsonarray=JSONArray.fromObject(a);
		for(int i=0;i<tiltl.length;i++)
		{
			cell=row.createCell(i);
			cell.setCellValue(tiltl[i]);
		}
		//System.out.println(rs.getString("nickname"));
		for(int i=1;rs.next();i++)
		{
				row=sheet.createRow(i);
				for(int j=0;j<tiltl.length;j++)
				{
					cell=row.createCell(j);
					if(j==0)
					{
						cell.setCellValue(i);
					}
					else if(j==tiltl.length-1)
					{
					
						String  flag="false";
						for(int k=0;k<jsonarray.size();k++)
						{														
							if(jsonarray.getString(k).equals(rs.getString(j)))
							{
								
								flag="true";								
							}							
						}
						if(flag.equals("true"))							
						{
							cell.setCellValue("关注");
						}
						else {
							cell.setCellValue("未关注");
						}
						
					}																
					else {
						cell.setCellValue(rs.getString(j));
						//System.out.println(rs.getString(j+1));
					}
					
				}
			}	
		String fileName = filename+".xls";
		fileName = URLEncoder.encode(fileName, "UTF-8");
		response.addHeader("Content-Disposition", "attachment;filename=" + fileName);
		OutputStream out = response.getOutputStream();
		wb.write(out);
		wb.close();
		out.close();
		return true;
	}
}
