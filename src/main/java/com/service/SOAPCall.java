package com.service;

import java.io.BufferedInputStream;
import java.io.BufferedReader;
import java.io.BufferedWriter;
import java.io.DataOutputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.io.ObjectOutputStream;
import java.io.OutputStream;
import java.io.OutputStreamWriter;
import java.io.UnsupportedEncodingException;
import java.net.HttpURLConnection;
import java.net.URL;
import java.net.URLConnection;
import java.net.URLEncoder;

import javax.net.ssl.HttpsURLConnection;
import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;

import org.json.JSONArray;
import org.json.JSONException;
import org.json.JSONObject;
import org.w3c.dom.Document;
import org.w3c.dom.NodeList;

import com.sun.jersey.core.util.Base64;

import javax.net.ssl.HostnameVerifier;
import javax.net.ssl.HttpsURLConnection;
import javax.net.ssl.SSLContext;
import javax.net.ssl.SSLSession;
import javax.net.ssl.TrustManager;
import javax.net.ssl.X509TrustManager;
import java.security.cert.X509Certificate;
import java.util.Collections;
import java.util.Comparator;
import java.util.HashMap;
import java.util.Map;
import java.util.Map.Entry;

public class SOAPCall {
	public static void main(String [] args) {
//		//String a="TemplateName=Template1&filename=Template1.docx&jsonstring={\"sdvsv\":\"vsdsv\"}";
//		JSONObject obj = new JSONObject();
		try {
			
//			String a="{\"requestDate\":\"2020-02-21 15:28:1\",\"createdby\":\"\",\"createddate\":\"2020-02-21 15:28:1\",\"createrequest\":{\"region\":[\"Gurgaon\",\"Pune\"],\"project\":[\"project2\",\"project3\"],\"requesttype\":\"new\",\"type\":\"Site Visit\",\"category\":\"AdHoc\",\"eventtype\":\"no\",\"campaigndate\":\"02/26/2020\",\"campaigntitle\":\"mnbhbjh\",\"campaignobjective\":\"jhvjh\",\"campaignbrief\":{\"What is the background to this Communication?\":\"vjhv\",\"What is this communication required? And Why?\":\"jhvjhv\",\"What actions do you wish your customer to take\":\"jhv\",\"Date, Time and Venue\":\"jhv\",\"Specific content relevant to the campaign\":\"jhv\",\"What is the communication objective?- Mandatory\":\"jhvjh\",\"What should we highlight in the communication and what should we stay away from?\":\"vjhv\",\"Target audience\":\"jhv\",\"Disclaimer\":\"jhv\"},\"eventbrief\":{},\"targetaudience\":\"\",\"status\":{}}}";
//			JSONObject aobj = new JSONObject(a); 
			//String res= new SOAPCall().callPostJSonModified("https://carrotrule.bluealgo.com:8083/drools/callrules/maitreyee.p@godrejproperties.com_Godrej_Agreement_Generation_Godrej_Agreement_RuleEngine/fire", obj);
		
			
// String res= new SOAPCall().callPostJSonModifiedWithAuth( "http://35.245.108.32:8088/request",  "ROUser",  "a",  aobj) ;
//	System.out.println("res&&&   "+res);

//	 String responsestr= new SOAPCall().callGetWithAuth( "http://35.245.108.32:8088/request",  "HOUser",  "a") ;
//		System.out.println("responsestr SaveCreateRequestServ"+responsestr);

//		String responsePdf="[{\"filename\":\"AgreementLetter_13-Mar-2020_19-17-30-597.pdf\",\"documentlink\":\"http://gpl.bluealgo.com:8082/Attachment/686f4838-a224-4841-abd1-8e21ced4af78.pdf\",\"creation_date\":\"Fri Mar 13 19:17:30 IST 2020\",\"created_by\":\"maitreyee.p@godrejproperties.com\"}]";	
//			
//		if(responsePdf.toString().indexOf("8082") != -1){
//			responsePdf.replace("http://gpl.bluealgo.com:8082", "https://gpl.bluealgo.com:8083" );
//			}
//			System.out.println("responsePdf  "+responsePdf);
//			
//			
			
			// String resp= new SOAPCall().callPostwithparam( "https://bluealgo.com:8092/NewMailDev/WhatsappSendMessage",  "to=+918689877203&bodyMsg=whatsapp body") ;


//			String requestURL="https://bluealgo.com:8092/NewMailDev/WhatsappSendMessage";
//			HashMap<String, String> postDataParams=new HashMap<>();
//		
//			postDataParams.put("to", "+918689877203"); // uk
//			postDataParams.put("bodyMsg", "Api Test whatapp body msg.");
//			
//			String data=new SOAPCall().performPostCall(requestURL, postDataParams);
//			System.out.println(data);
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}
	
	
	

	
	 public  String callGet(String url) throws Exception {
		 
		 
		 ignoreHttps( url);
         StringBuffer userresult=null;
	//String url = "http://35.221.160.146:8180/UserValidation/services/UserValidation/raveUserExistence?userId=doctiger100@gmail.com";
	try {
	URL obj = new URL(url);
	HttpURLConnection con = (HttpURLConnection) obj.openConnection();
	// optional default is GET
	con.setRequestMethod("GET");
	//con.setRequestProperty("Content-Type", "application/xml;charset=UTF-8");
	int responseCode = con.getResponseCode();
	System.out.println("\nSending 'GET' request to URL : " + url);
	System.out.println("Response Code : " + responseCode);
	BufferedReader in = new BufferedReader(new InputStreamReader(con.getInputStream()));
		String inputLine;
		userresult = new StringBuffer();

		while ((inputLine = in.readLine()) != null) {
			userresult.append(inputLine);
		}
		in.close();

		//System.out.println(userresult.toString());
		
		
	
 System.out.println("userresult  "+userresult);
return userresult.toString();
	}catch(Exception e) {
		//e.printStackTrace();
		return e.getMessage();
	}
	
	
}

	
	
	
	public  String callPostService(String urlStr, String  param) {
		
		 ignoreHttps( urlStr);

String userresult="";
		try {
			  // TODO Auto-generated method stub
		//	  String json = "{\"Fetcher_Id\":\"vivek123@bizlem.com\",\"Template_Name\":\"templatetest1\",\"Subject\":[\"sdsvd\"],\"Body\":[],\"Attachment\":[],\"Attachment_Name\":[],\"match\":\"1\",\"Alert\":\"\"}";
			  String urlParameters =param;
			  urlParameters = urlParameters.replace(" ", "%20");
			//  URL url = new URL("http://35.186.166.22:8082/portal/servlet/service/MapMailSolrServ.tem1");
			  URL url = new URL(urlStr);

			  URLConnection conn = url.openConnection();
			  conn.setDoOutput(true);
			  OutputStreamWriter writer = new OutputStreamWriter(conn.getOutputStream());

			  writer.write(urlParameters);
			  writer.flush();
InputStream inputXml= conn.getInputStream();

DocumentBuilderFactory dbFactory = DocumentBuilderFactory.newInstance();
DocumentBuilder dBuilder = dbFactory.newDocumentBuilder();
Document doc = dBuilder.parse(inputXml);
doc.getDocumentElement().normalize();
NodeList nList1 = doc.getElementsByTagName("ns:DocGenerationResponse");
org.w3c.dom.Node nNode = nList1.item(0);
org.w3c.dom.Element eElement = (org.w3c.dom.Element) nNode;
 userresult = eElement.getElementsByTagName("ns:return").item(0).getTextContent();

System.out.println("userresult  "+userresult);

			//  String line;
			//  BufferedReader reader = new BufferedReader(new InputStreamReader(conn.getInputStream()));

//			  while ((line = reader.readLine()) != null) {
//			   //   System.out.println(line);
//				  res=  res+line;
//			  }
//			  reader.close();         
			  
writer.close();
	  	  }catch (Exception e) {
			   // TODO: handle exception
			   System.out.println("error :: "+e.getMessage());
			  }
		return userresult;
    }


public String callPostJSon(String urlstr, JSONObject Obj) {
	StringBuffer response =null;
		  int responseCode = 0;
		  String urlParameters = "";
		  try {
				 ignoreHttps( urlstr);


		   URL url = new URL(urlstr);
		   HttpURLConnection con = (HttpURLConnection) url.openConnection();
		   con.setRequestMethod("POST");

		   con.setRequestProperty("Content-Type", "application/json");
		   con.setRequestProperty("Accept-Charset", "UTF-8");
		  
		   con.setDoOutput(true);
		   DataOutputStream wr = new DataOutputStream(con.getOutputStream());
		   wr.writeBytes(Obj.toString());
		   wr.flush();
		   wr.close();

		   responseCode = con.getResponseCode();

		   BufferedReader in = new BufferedReader(new InputStreamReader(con.getInputStream()));
		   String inputLine;
		    response = new StringBuffer();

		   while ((inputLine = in.readLine()) != null) {
		    response.append(inputLine);
		   }
		   in.close();

		   System.out.println("post op "+response.toString());
		  }
		  catch (Exception e) {
		return   e.getMessage();
		  }
		  return response.toString();

		 }


public String callPostJSonModified(String urlstr, JSONObject Obj) {
	
		StringBuffer response =null;
		int responseCode = 0;
		String urlParameters = "";
		try {
            ignoreHttps(urlstr);
		URL url = new URL(urlstr);
		HttpURLConnection con = (HttpURLConnection) url.openConnection();
		con.setRequestMethod("POST");

		con.setRequestProperty("Content-Type", "application/json;charset=UTF-8");
		con.setRequestProperty("Accept-Charset", "UTF-8");

		con.setDoOutput(true);
		// DataOutputStream wr = new DataOutputStream(con.getOutputStream());
		// wr.writeBytes(Obj.toString());
		// wr.flush();
		// wr.close();
//System.out.println("Obj "+Obj);
		DataOutputStream wr = new DataOutputStream(con.getOutputStream());
		BufferedWriter writer = new BufferedWriter(new OutputStreamWriter(wr, "UTF-8"));
		writer.write(Obj.toString());
		writer.close();
		wr.close();

		responseCode = con.getResponseCode();

		BufferedReader in = new BufferedReader(new InputStreamReader(con.getInputStream()));
		String inputLine;
		response = new StringBuffer();

		while ((inputLine = in.readLine()) != null) {
		response.append(inputLine);
		}
		in.close();

		System.out.println(response.toString());
		}
		catch (Exception e) {
		return e.getMessage();
		}
		return response.toString();

		}

		




public String callGetWithAuth(String urlstr, String username, String password) {
	StringBuffer response =null;
		  int responseCode = 0;
		  String urlParameters = "";
		  try {
			  ignoreHttps(urlstr);
		   URL url = new URL(urlstr);
		   HttpURLConnection con = (HttpURLConnection) url.openConnection();
		   con.setRequestMethod("GET");

		//   BASE64Encoder enc = new sun.misc.BASE64Encoder();
		      String userpassword = username + ":" + password;
		    //  String encodedAuthorization = enc.encode( userpassword.getBytes() );
		         byte[] encodedAuthorization = Base64.encode(userpassword);
		        // byte[] authEncBytes = Base64.encodeBase64(authString.getBytes());
					String authStringEnc = new String(encodedAuthorization);
					

		      con.setRequestProperty("Authorization", "Basic "+
		    		  authStringEnc);
		   con.setRequestProperty("Content-Type", "application/json");

		  
		   con.setDoOutput(true);
		   

		   responseCode = con.getResponseCode();
		   System.out.println("responseCode  "+responseCode);

		   BufferedReader in = new BufferedReader(new InputStreamReader(con.getInputStream()));
		   String inputLine;
		    response = new StringBuffer();

		   while ((inputLine = in.readLine()) != null) {
		    response.append(inputLine);
		   }
		   in.close();

		   System.out.println("responseCode **** "+response.toString());
		  }
		  catch (Exception e) {
		return   e.getMessage();
		  }
		  return response.toString();

		 }



public String callPostJSonModifiedWithAuth(String urlstr, String username, String password, JSONObject Obj) {
	
	StringBuffer response =null;
	int responseCode = 0;
	String urlParameters = "";
	try {
        ignoreHttps(urlstr);
	URL url = new URL(urlstr);
	HttpURLConnection con = (HttpURLConnection) url.openConnection();
	con.setRequestMethod("POST");
	
	//   BASE64Encoder enc = new sun.misc.BASE64Encoder();
	      String userpassword = username + ":" + password;
	    //  String encodedAuthorization = enc.encode( userpassword.getBytes() );
	         byte[] encodedAuthorization = Base64.encode(userpassword);
	        // byte[] authEncBytes = Base64.encodeBase64(authString.getBytes());
				String authStringEnc = new String(encodedAuthorization);
	      con.setRequestProperty("Authorization", "Basic "+
	    		  authStringEnc);

	con.setRequestProperty("Content-Type", "application/json;charset=UTF-8");
	con.setRequestProperty("Accept-Charset", "UTF-8");

	con.setDoOutput(true);
	// DataOutputStream wr = new DataOutputStream(con.getOutputStream());
	// wr.writeBytes(Obj.toString());
	// wr.flush();
	// wr.close();
//System.out.println("Obj "+Obj);
	DataOutputStream wr = new DataOutputStream(con.getOutputStream());
	BufferedWriter writer = new BufferedWriter(new OutputStreamWriter(wr, "UTF-8"));
	writer.write(Obj.toString());
	writer.close();
	wr.close();

	responseCode = con.getResponseCode();

	BufferedReader in = new BufferedReader(new InputStreamReader(con.getInputStream()));
	String inputLine;
	response = new StringBuffer();

	while ((inputLine = in.readLine()) != null) {
	response.append(inputLine);
	}
	in.close();

	System.out.println(response.toString());
	}
	catch (Exception e) {
	return e.getMessage();
	}
	return response.toString();

	}

	




public String callPutWithAuth(String urlstr, String username, String password) {
	StringBuffer response =null;
		  int responseCode = 0;
		  String urlParameters = "";
		  try {
			  ignoreHttps(urlstr);

		   URL url = new URL(urlstr);
		   HttpURLConnection con = (HttpURLConnection) url.openConnection();
		   con.setRequestMethod("PUT");

		//   BASE64Encoder enc = new sun.misc.BASE64Encoder();
		      String userpassword = username + ":" + password;
		    //  String encodedAuthorization = enc.encode( userpassword.getBytes() );
		         byte[] encodedAuthorization = Base64.encode(userpassword);
		        // byte[] authEncBytes = Base64.encodeBase64(authString.getBytes());
					String authStringEnc = new String(encodedAuthorization);
					

		      con.setRequestProperty("Authorization", "Basic "+
		    		  authStringEnc);
		   con.setRequestProperty("Content-Type", "application/json");

		  
		   con.setDoOutput(true);
		   

		   responseCode = con.getResponseCode();

//		   BufferedReader in = new BufferedReader(new InputStreamReader(con.getInputStream()));
//		   String inputLine;
//		    response = new StringBuffer();
//
//		   while ((inputLine = in.readLine()) != null) {
//		    response.append(inputLine);
//		   }
//		   in.close();
//
//		   System.out.println(response.toString());
		  }
		  catch (Exception e) {
		return   e.getMessage();
		  }
		  return Integer.toString(responseCode);

		 }


      
       
       
       
       public void ignoreHttps(String urlstring){
    	   try{
    	   if(urlstring.indexOf("https:") != -1){
    	   TrustManager[] trustAllCerts = new TrustManager[] {
    	   new X509TrustManager() {
    	   public java.security.cert.X509Certificate[] getAcceptedIssuers() {
    	   return null;
    	   }

    	   public void checkClientTrusted(X509Certificate[] certs, String authType) { }

    	   public void checkServerTrusted(X509Certificate[] certs, String authType) { }

    	   }
    	   };


    	   try {
    	   SSLContext sc = SSLContext.getInstance("SSL");
    	   sc.init(null, trustAllCerts, new java.security.SecureRandom());
    	   HttpsURLConnection.setDefaultSSLSocketFactory(sc.getSocketFactory());

    	   } catch (Exception e1) {
    	   // TODO Auto-generated catch block
    	   e1.printStackTrace();
    	   }

    	   // Create all-trusting host name verifier
    	   HostnameVerifier allHostsValid = new HostnameVerifier() {
    	   public boolean verify(String hostname, SSLSession session) {
    	   return true;
    	   }
    	   };
    	   // Install the all-trusting host verifier
    	   HttpsURLConnection.setDefaultHostnameVerifier(allHostsValid);
    	   /*
    	   * end of the fix
    	   */
    	   }
    	   }catch(Exception e){
            e.printStackTrace();
    	   }
    	   }
       
       
       
       public String callGetAccess_token(String urlstr, String access_token) {
    	   StringBuffer response = new StringBuffer();
    	   try {
    			  ignoreHttps(urlstr);

    	   URL url = new URL(urlstr);
    	   HttpURLConnection conn = (HttpURLConnection) url.openConnection();

    	   conn.setRequestProperty("Authorization","Bearer "+" "+access_token);

    	   conn.setRequestProperty("Content-Type","application/json");
    	   conn.setRequestMethod("GET");


    	   BufferedReader in = new BufferedReader(new InputStreamReader(conn.getInputStream()));
    	   String output;

    	   while ((output = in.readLine()) != null) {
    	   response.append(output);
    	   }

    	   in.close();
    	   // printing result from response
    	   System.out.println("Response:-" + response.toString());

    	   
       }catch(Exception e) {
    	   e.printStackTrace();
       }
    	   return response.toString();
       }
       
       
       
       
       public static boolean isJSONValid(String test) {
   		try {
   			new JSONObject(test);
   		} catch (JSONException ex) {
   			
   			try {
   				new JSONArray(test);
   			} catch (JSONException ex1) {
   				return false;
   			}
   		}
   		return true;
   	}
   	
       

public String callPostwithparam(String urlstr, String paramstr) {
	
	
		StringBuffer response =null;
		int responseCode = 0;
		String urlParameters = "";
		try {
			
			
            ignoreHttps(urlstr);

			//byte[] postData       = urlParameters.getBytes( StandardCharsets.UTF_8 );
			//int    postDataLength = postData.length;
			URL    url            = new URL( urlstr );
			HttpURLConnection conn= (HttpURLConnection) url.openConnection();           
			conn.setDoOutput( true );
			conn.setInstanceFollowRedirects( false );
			conn.setRequestMethod( "POST" );
			conn.setRequestProperty( "Content-Type", "application/x-www-form-urlencoded"); 
			conn.setRequestProperty( "charset", "utf-8");
			//conn.setRequestProperty( "Content-Length", Integer.toString( postDataLength ));
			conn.setUseCaches( false );
			
			
			DataOutputStream wr = new DataOutputStream(conn.getOutputStream());
			BufferedWriter writer = new BufferedWriter(new OutputStreamWriter(wr, "UTF-8"));
			writer.write(paramstr);
			writer.close();
			wr.close();
		responseCode = conn.getResponseCode();

		BufferedReader in = new BufferedReader(new InputStreamReader(conn.getInputStream()));
		String inputLine;
		response = new StringBuffer();
		while ((inputLine = in.readLine()) != null) {
		response.append(inputLine);
		}
		in.close();

		System.out.println(response.toString());
		}
		catch (Exception e) {
		return e.getMessage();
		}
		return response.toString();

		}

public  String  performPostCall(String requestURL,HashMap<String, String> postDataParams){

    URL url;
    String response = "";
    try {
        url = new URL(requestURL);

        ignoreHttps(requestURL);

        HttpURLConnection conn = (HttpURLConnection) url.openConnection();
        conn.setReadTimeout(15000);
        conn.setConnectTimeout(15000);
        conn.setRequestMethod("POST");
        conn.setDoInput(true);
        conn.setDoOutput(true);

        OutputStream os = conn.getOutputStream();
        BufferedWriter writer = new BufferedWriter(new OutputStreamWriter(os, "UTF-8"));
        writer.write(getPostDataString(postDataParams));

        writer.flush();
        writer.close();
        os.close();
        int responseCode=conn.getResponseCode();

        if (responseCode == HttpsURLConnection.HTTP_OK) {
            String line;
            BufferedReader br=new BufferedReader(new InputStreamReader(conn.getInputStream()));
            while ((line=br.readLine()) != null) {
                response+=line;
            }
        }
        else {
            response="";

        }
    } catch (Exception e) {
        e.printStackTrace();
    }

    return response;
}

private static String getPostDataString(HashMap<String, String> params) throws UnsupportedEncodingException{
    StringBuilder result = new StringBuilder();
    boolean first = true;
    for(Map.Entry<String, String> entry : params.entrySet()){
        if (first)
            first = false;
        else
            result.append("&");

        result.append(URLEncoder.encode(entry.getKey(), "UTF-8"));
        result.append("=");
        result.append(URLEncoder.encode(entry.getValue(), "UTF-8"));
    }

    return result.toString();
}	


public static boolean isNullString(String p_text) {
	if (p_text != null && p_text.trim().length() > 0 && !"null".equalsIgnoreCase(p_text.trim())) {
		return false;
	} else {
		return true;
	}
}

public String saveTemplate(String fileurl, String savepath, String filename) {
	String resp = "";

	try {
		
		ignoreHttps(fileurl);
		
		fileurl=fileurl.replace(" ", "%20");
		BufferedInputStream in = new BufferedInputStream(new URL(fileurl).openStream());
		FileOutputStream fileOutputStream = new FileOutputStream(savepath + filename);
		byte dataBuffer[] = new byte[1024];
		int bytesRead;
		while ((bytesRead = in.read(dataBuffer, 0, 1024)) != -1) {
			fileOutputStream.write(dataBuffer, 0, bytesRead);
		}
		resp = "success";
		System.out.println("files saved");
	} catch (Exception e) {
		// TODO: handle exception
		resp = e.getMessage();
		System.out.println("e :: " + e);
	} finally {

	}
	return resp;
}

       
       
       
       
       
       
}
