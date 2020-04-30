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

import com.service.PPTWork;
import com.service.SOAPCall;

/**
 * Servlet implementation class PPTModificationServ
 */
public class PPTModificationServ extends HttpServlet {
	private static final long serialVersionUID = 1L;
	static ResourceBundle bundleststic = ResourceBundle.getBundle("config_PPTExcel");

    /**
     * @see HttpServlet#HttpServlet()
     */
    public PPTModificationServ() {
        super();
        // TODO Auto-generated constructor stub
    }

	/**
	 * @see HttpServlet#doGet(HttpServletRequest request, HttpServletResponse response)
	 */
	protected void doGet(HttpServletRequest request, HttpServletResponse response) throws ServletException, IOException {
		// TODO Auto-generated method stub
		response.getWriter().append("Served at: ").append(request.getContextPath());
	}

	/**
	 * @see HttpServlet#doPost(HttpServletRequest request, HttpServletResponse response)
	 */
	protected void doPost(HttpServletRequest request, HttpServletResponse response) throws ServletException, IOException {
		// TODO Auto-generated method stub
		PrintWriter out = response.getWriter();
		response.setCharacterEncoding("UTF-8");
		response.setHeader("Content-Type", "text/html,charset=UTF-8");
		String TemplateUrl="";
		String accesstype="public";
		String ppttemplatepath="";
		String PPTSavepath="";
		String filename="";
		String result="";
		String PPTReturnUrl="";
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
	
ppttemplatepath=bundleststic.getString("uploaded_templates_path");
PPTSavepath=bundleststic.getString("DocGenServerPPTFilePath");
if(TemplateUrl.lastIndexOf("/")!=-1) {
 filename=TemplateUrl.substring(TemplateUrl.lastIndexOf("/")+1);
}


Random rand = new Random(); 
int rand_int2 = rand.nextInt(1000); 

PPTReturnUrl=bundleststic.getString("PPTReturnUrlpath")+rand_int2+filename;
String tempasaveresult= soapcall.saveTemplate(TemplateUrl, ppttemplatepath, filename);
if(tempasaveresult.equalsIgnoreCase("success")) {
	
	 result= new PPTWork().parsePPT(ppttemplatepath+filename, resultobj, PPTSavepath+rand_int2+filename);
	// out.print(result);
	 
	 
	 
	 
	 out.print(PPTReturnUrl);

}

	


		}catch (Exception e) {
			// TODO: handle exception
			out.println(e.getMessage().toString());
e.printStackTrace();
		}
	
	
		
		
	}

}
