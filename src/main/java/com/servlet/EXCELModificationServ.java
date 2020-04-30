package com.servlet;

import java.io.BufferedInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.PrintWriter;
import java.util.Random;
import java.util.ResourceBundle;

import javax.servlet.ServletException;
import javax.servlet.http.HttpServlet;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import org.json.JSONObject;

import com.service.EXCELWork;
import com.service.PPTWork;
import com.service.SOAPCall;

/**
 * Servlet implementation class EXCELModificationServ
 */
public class EXCELModificationServ extends HttpServlet {
	private static final long serialVersionUID = 1L;
	static ResourceBundle bundleststic = ResourceBundle.getBundle("config_PPTExcel");

    
    public EXCELModificationServ() {
        super();
        // TODO Auto-generated constructor stub
    }

	
	protected void doGet(HttpServletRequest request, HttpServletResponse response) throws ServletException, IOException {
		// TODO Auto-generated method stub
		response.getWriter().append("Served at: ").append(request.getContextPath());
	}

	
	protected void doPost(HttpServletRequest request, HttpServletResponse response) throws ServletException, IOException {
		// TODO Auto-generated method stub
		PrintWriter out = response.getWriter();
		response.setCharacterEncoding("UTF-8");
		response.setHeader("Content-Type", "text/html,charset=UTF-8");
		String TemplateUrl="";
		String accesstype="public";
		String EXCELtemplatepath="";
		String EXCELSavepath="";
		String filename="";
		String result="";
		String EXCELReturnUrl="";
		try {
			SOAPCall soapcall= new SOAPCall();
			BufferedInputStream bis = new BufferedInputStream(request.getInputStream());
			ByteArrayOutputStream buf = new ByteArrayOutputStream();
			int result1 = bis.read();
			while (result1 != -1) {
				buf.write((byte) result1);
				result1 = bis.read();
			}
			// StandardCharsets.UTF_8.name() > JDK 7
			String res = buf.toString("UTF-8");			
			System.out.println("res: " + res);
			JSONObject resultobj = new JSONObject(res);
if(resultobj.has("TemplateUrl")){TemplateUrl=resultobj.getString("TemplateUrl");}
if(resultobj.has("accesstype")){accesstype=resultobj.getString("accesstype");}
	
EXCELtemplatepath=bundleststic.getString("uploaded_templates_path");
EXCELSavepath=bundleststic.getString("DocGenServerEXCELFilePath");
if(TemplateUrl.lastIndexOf("/")!=-1) {
 filename=TemplateUrl.substring(TemplateUrl.lastIndexOf("/")+1);
}


Random rand = new Random(); 
int rand_int2 = rand.nextInt(1000); 

EXCELReturnUrl=bundleststic.getString("EXCELReturnUrlpath")+rand_int2+filename;
String tempasaveresult= soapcall.saveTemplate(TemplateUrl, EXCELtemplatepath, filename);
if(tempasaveresult.equalsIgnoreCase("success")) {
	
	 result= new EXCELWork().parseXLSX(EXCELtemplatepath+filename, resultobj, EXCELSavepath+rand_int2+filename);
	// out.print(result);
	
	
	 out.print(EXCELReturnUrl);

}

		}catch (Exception e) {
			// TODO: handle exception
			out.println(e.getMessage().toString());
e.printStackTrace();
		}		
	}

}
