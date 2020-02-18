package ppt;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.io.PrintWriter;
import java.nio.charset.Charset;
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import javax.servlet.ServletException;
import javax.servlet.annotation.WebServlet;
import javax.servlet.http.HttpServlet;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import org.apache.commons.fileupload.DiskFileUpload;
import org.apache.commons.fileupload.FileItem;
import org.apache.commons.fileupload.FileUploadException;
import org.apache.commons.fileupload.disk.DiskFileItemFactory;
import org.apache.commons.fileupload.servlet.ServletFileUpload;



import net.sf.json.JSON;


/**
 * Servlet implementation class url
 */
@WebServlet("/url")
public class url extends HttpServlet {
	private static final long serialVersionUID = 1L;
       
    /**
     * @see HttpServlet#HttpServlet()
     */
    public url() {
        super();
        // TODO Auto-generated constructor stub
    }

	/**
	 * @see HttpServlet#doGet(HttpServletRequest request, HttpServletResponse response)
	 */
	protected void doGet(HttpServletRequest request, HttpServletResponse response) throws ServletException, IOException {
		// TODO Auto-generated method stub
		
	}

	/**
	 * @see HttpServlet#doPost(HttpServletRequest request, HttpServletResponse response)
	 */
	protected void doPost(HttpServletRequest request, HttpServletResponse response) throws ServletException, IOException {
		
	request.setCharacterEncoding("utf-8");
	response.setContentType("text/html;charset=utf-8");
	PrintWriter out=response.getWriter();
	Pattern p = Pattern.compile(".*pptx.*");
	Matcher m=null;
	  DiskFileItemFactory factory = new DiskFileItemFactory();
	  ServletFileUpload upload = new ServletFileUpload(factory);   
      upload.setHeaderEncoding("UTF-8");  
      List items;
      Map params = new HashMap();
	try {
		items = upload.parseRequest(request);
		
		 for(Object object:items){  
	          FileItem fileItem = (FileItem) object;   
	           InputStream is=fileItem.getInputStream();
	           Matcher l=p.matcher(fileItem.getFieldName());
	           if(l.matches()) {
	        	   String path=request.getServletContext().getRealPath("/WEB-INF/datapptx/"+fileItem.getFieldName());
	        	   File tempFile=new File(path);
	        	   FileOutputStream fos=new FileOutputStream(tempFile);
	        	   byte[]b=new byte[1024];
	        	   int n;
	        	   while ((n=is.read(b)) !=-1) {
					fos.write(b,0,n);
					
				}
	        	   fos.close();
	        	   is.close();
	           }
	          
	           
	      }   
		
		 out.println("±£´æ³É¹¦");
	} catch (FileUploadException e) {
		// TODO Auto-generated catch block
		e.printStackTrace();
	}  
        
     

    out.close();
	
	
	
		}
		

}
