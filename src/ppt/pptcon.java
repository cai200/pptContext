package ppt;

import java.io.BufferedInputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;

import javax.servlet.ServletException;
import javax.servlet.annotation.WebServlet;
import javax.servlet.http.HttpServlet;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFSlide;

/**
 * Servlet implementation class pptcon
 */
@WebServlet("/pptcon")
public class pptcon extends HttpServlet {
	private static final long serialVersionUID = 1L;

    /**
     * Default constructor. 
     */
    public pptcon() {
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
				String file1 = "E:/test1.pptx";
			    String file2 = "E:/test2.pptx";
			 
			    FileInputStream is = new FileInputStream(file1);
			    XMLSlideShow src = new XMLSlideShow(is);
			    is.close();
			    FileInputStream is2 = new FileInputStream(file2);
			    XMLSlideShow src2 = new XMLSlideShow(is2);
			    is2.close();
			 
			    XMLSlideShow ppt = new XMLSlideShow();
			    for (XSLFSlide slide : src.getSlides()) {
			        XSLFSlide slide1 = ppt.createSlide(slide.getSlideLayout());
			        slide1.importContent(slide);
			    }
			    for (XSLFSlide slide : src2.getSlides()) {
			        XSLFSlide slide1 = ppt.createSlide(slide.getSlideLayout());
			        slide1.importContent(slide);
			    }
			    String path = request.getServletContext().getRealPath("/WEB-INF/out.pptx");
			    FileOutputStream out = new FileOutputStream(path);
			    
			    ppt.write(out);
			    save(request,response,path,"out","pptx");
			    out.close();
			    System.out.println("merge successfully");
		    }
	
	
	public void save(HttpServletRequest request,HttpServletResponse response,String path,String name,String type) throws IOException {	 
		FileInputStream fis = null;
		BufferedInputStream bis = null;
		OutputStream os = null;
	   try {
	//获取下载文件所在路径
//	    String path = request.getServletContext().getRealPath("/WEB-INF/data/"+name);
		File file = new File(path);
		//判断文件是否存在
	   if(file.exists()){
	//且仅当此对象抽象路径名表示的文件或目录存在时，返回true
	     response.setContentType("application/"+name+"."+type);
		//控制下载文件的名字
	     response.addHeader("Content-Disposition", "attachment;filename="+name+"."+type);   
		byte buf[] = new byte[1024];
		fis = new FileInputStream(file);
		bis = new BufferedInputStream(fis);
		os = response.getOutputStream();
		int i = bis.read(buf);
		while(i!=-1){
			os.write(buf,0,i);
			i = bis.read(buf);

	}
		
	       }else{
	    	   System.out.println("您下载的资源已经不存在！！");

	}
		

	} catch (Exception e) {
	    e.printStackTrace();
	} finally {
		os.close();
		bis.close();
		fis.close();
	}
	}
	
	
	
		}



