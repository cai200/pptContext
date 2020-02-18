package ppt;

import java.io.BufferedInputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.Iterator;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import javax.servlet.ServletException;
import javax.servlet.annotation.WebServlet;
import javax.servlet.http.HttpServlet;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFSlide;

import net.sf.json.JSONArray;
import net.sf.json.JSONObject;

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
		response.setContentType("text/html;charset=utf-8");
		//�������ݱ����ʽ����ֹ��������
		request.setCharacterEncoding("utf-8");
		 String fileName = request.getParameter("fileName");
	//ǰ̨����jsonObject=[{"data":"1"url":"d:/text1.ppt"}{"data":"2"url":"d:/text2.ppt"},..]
		 JSONArray jsonObject = JSONArray.fromObject(fileName);	
		 System.out.print(jsonObject);
		// TODO Auto-generated method stub
		/* 
		 * ����ǰ̨�ϴ���N��ppt,���Բ�����ÿ��������һ���������Խ����������ʹ�ñ��������õ���ֵ��
		 * Ȼ������ٴα���д�룬������±�պ���ͬ�ġ���ʱһһ��Ӧ
		 *  
		 *  */
		String file[]=new String[20];
		FileInputStream is[] = new FileInputStream[20];
		  XMLSlideShow src[] = new XMLSlideShow[20];
		  String path1[]=new String[20];
		  File filepath[]=null;
		  Pattern p = Pattern.compile(".*pptx.*");
			Matcher m=null;
		/*
		 * String file1 = "E:/test1.pptx"; String file2 = "E:/test2.pptx";
		 */
		  for (int i = 0; i < jsonObject.size(); i++) {
			  JSONObject Object = JSONObject.fromObject(jsonObject.get(i));	
			  Iterator<String> iterator1 =Object.keys();
	            
	            while(iterator1.hasNext()){
	            	 String key1 = iterator1.next();
	                 String value1 = Object.getString(key1);
	                 //ǰ̨Ҫǿ������1.2.3...,�Ͳ���Ҫ�ٴ��ж�
	                 Matcher l=p.matcher(Object.getString(key1));
	  	           if(l.matches()) {
	            		//urlΪǰ̨���͹����ļ�ֵ
						file[i]=Object.getString("name");
						path1[i]=request.getServletContext().getRealPath("/WEB-INF/datapptx/"+file[i]);
						 //filepath[i]=new File(path1[i]);
						is[i]= new FileInputStream(path1[i]);
						
						 src[i] = new XMLSlideShow(is[i]);
						 is[i].close();
	            	}
	            }
		}
		/*
		 * FileInputStream is = new FileInputStream(file1); XMLSlideShow src = new
		 * XMLSlideShow(is);
		 * 
		 * FileInputStream is2 = new FileInputStream(file2); XMLSlideShow src2 = new
		 * XMLSlideShow(is2);
		 */
		  XMLSlideShow ppt = new XMLSlideShow();
		  for (int i = 0; i < jsonObject.size(); i++) { 
			 
			    
			    for (XSLFSlide slide : src[i].getSlides()) {
			        XSLFSlide slide1 = ppt.createSlide(slide.getSlideLayout());
			        slide1.importContent(slide);
			    }
		  }
			    /*for (XSLFSlide slide : src2.getSlides()) {
			        XSLFSlide slide1 = ppt.createSlide(slide.getSlideLayout());
			        slide1.importContent(slide);
			    }*/
			    String path = request.getServletContext().getRealPath("/WEB-INF/out.pptx");
			    FileOutputStream out = new FileOutputStream(path);
			    
			    ppt.write(out);
			    save(request,response,path,"out","pptx");
			    out.close();
			    //���غ�ɾ��ǰ̨�ϴ���ʱ������datapptx��pptx�ļ�
			    del(request,jsonObject);
			    System.out.println("merge successfully");
		    }
	
	
	public void save(HttpServletRequest request,HttpServletResponse response,String path,String name,String type) throws IOException {	 
		FileInputStream fis = null;
		BufferedInputStream bis = null;
		OutputStream os = null;
	   try {
	//��ȡ�����ļ�����·��
//	    String path = request.getServletContext().getRealPath("/WEB-INF/data/"+name);
		File file = new File(path);
		//�ж��ļ��Ƿ����
	   if(file.exists()){
	//�ҽ����˶������·������ʾ���ļ���Ŀ¼����ʱ������true
	     response.setContentType("application/"+name+"."+type);
		//���������ļ�������
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
	    	   System.out.println("�����ص���Դ�Ѿ������ڣ���");

	}
		

	} catch (Exception e) {
	    e.printStackTrace();
	} finally {
		os.close();
		bis.close();
		fis.close();
	}
	}
	
	
	
	
	public void del(HttpServletRequest request,JSONArray jsonObject) {
		 String path1[]=new String[20];
		 File file=null;
		 Pattern p = Pattern.compile(".*pptx.*");
			Matcher m=null;
		 for (int i = 0; i < jsonObject.size(); i++) {
			  JSONObject Object = JSONObject.fromObject(jsonObject.get(i));	
			  Iterator<String> iterator1 =Object.keys();
	            
	            while(iterator1.hasNext()){
	            	
	            	 String key1 = iterator1.next();
	                 String value1 = Object.getString(key1);
	                 //ǰ̨Ҫǿ������1.2.3...,�Ͳ���Ҫ�ٴ��ж�
	                 Matcher l=p.matcher(Object.getString(key1));
	  	           if(l.matches()) {
	  	        	 path1[i]=request.getServletContext().getRealPath("/WEB-INF/datapptx/"+value1);
	  	        	 file=new File(path1[i]);
	  	        	 file.delete();
	  	           }
	            	
	            }}
	}
	
	
	
	
	
	
		}



