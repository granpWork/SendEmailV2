package com.ltg;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Date;
import java.util.Iterator;
import java.util.List;
import java.util.Properties;

import javax.mail.Authenticator;
import javax.mail.Message;
import javax.mail.Multipart;
import javax.mail.PasswordAuthentication;
import javax.mail.Session;
import javax.mail.Transport;
import javax.mail.Message.RecipientType;
import javax.mail.internet.InternetAddress;
import javax.mail.internet.MimeBodyPart;
import javax.mail.internet.MimeMessage;
import javax.mail.internet.MimeMultipart;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Main {
	
	static String filePath;
	
	static String hostMail;
	static String portMail;
	static String usernameMail;
	static String passwordMail;
	static String emailAddressFrom;

	public static void main(String[] args) {
		
		final String readFile = "C:\\Users\\emylyn.audemard\\Documents\\sendEmailJava\\emails.xlsx";
		final String readProp = "C:\\Users\\emylyn.audemard\\Documents\\sendEmailJava\\smptconfig.properties";
		
		System.out.println("Running SendEmailV2");
		System.out.println("File Validation....");
		if(!checkFile(".properties", readProp) || !checkFile(".xlsx", readFile)) {
			System.exit(0);
		}else {
			System.out.println("All files are valid.");
			System.out.println("");
		}

		setReadProp(readProp);
		setFilePath(readFile);
		
		for(String email : getEmailsInList()) {
			sendEmail(email);
		}
		
		System.out.println("Done Sending");
		System.exit(0);

	}

	public static String getFilePath() {
		return filePath;
	}

	public static void setFilePath(String fp) {
		filePath = fp;
	}
	
	public static String getHostMail() {
		return hostMail;
	}

	public static void setHostMail(String hostMail) {
		Main.hostMail = hostMail;
	}

	public static String getPortMail() {
		return portMail;
	}

	public static void setPortMail(String portMail) {
		Main.portMail = portMail;
	}

	public static String getUsernameMail() {
		return usernameMail;
	}

	public static void setUsernameMail(String usernameMail) {
		Main.usernameMail = usernameMail;
	}

	public static String getPasswordMail() {
		return passwordMail;
	}

	public static void setPasswordMail(String passwordMail) {
		Main.passwordMail = passwordMail;
	}

	public static String getEmailAddressFrom() {
		return emailAddressFrom;
	}

	public static void setEmailAddressFrom(String emailAddressFrom) {
		Main.emailAddressFrom = emailAddressFrom;
	}
	
	public static void setReadProp(String readProp) {

		try {
			InputStream input = new FileInputStream(readProp);
			Properties prop = new Properties();
			
			prop.load(input);
			
            // get the property value and print it out
			setUsernameMail(prop.getProperty("mail.username"));
			setPasswordMail(prop.getProperty("mail.password"));
			setHostMail(prop.getProperty("mail.smtp.host"));
			setPortMail(prop.getProperty("mail.smtp.port"));
			setEmailAddressFrom(prop.getProperty("mail.email.address.from"));
            
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		
	}

	public static boolean checkFile(String fileformat, String file) {
			File f = new File(file);
			if(f.isFile() && !f.isDirectory()) { 
				
				String filename = f.getName().toLowerCase();
				
				if(!filename.endsWith(fileformat)) {
					System.out.println(file + " is not valid excel format.");
					
					return false;
				}
			}else {
				System.out.println(file + " does not exist.");
				
				return false;
			}
			
			return true;
		}
		
	public static List<String> getEmailsInList(){
			List<String> list = new ArrayList<String>(); 

			try {
				
				FileInputStream fis = new FileInputStream(new File(getFilePath()));
				XSSFWorkbook workbook = new XSSFWorkbook(fis);
				XSSFSheet spreadsheet = workbook.getSheetAt(0);

				  
				  Iterator < Row >  rowIterator = spreadsheet.iterator();
				  
				  while (rowIterator.hasNext()) {
					  Row row = rowIterator.next();
					  Iterator<Cell> cellIterator = row.cellIterator();
					  
					  while (cellIterator.hasNext()) {
				    		Cell cell = cellIterator.next();
				    		
				    		switch (cell.getCellType()) {
				               case STRING:
					    			list.add(cell.getStringCellValue().toString());
				                  break;
				               case BLANK:
						              break;
							default:
								break;
				            }
					  }
				  }
				  
				  workbook.close();
				  
			} catch (FileNotFoundException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			} catch (IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
			
			System.out.println("Total Email Address: "+list.size());
			
			return list;
	}
		
	@SuppressWarnings("unused")
	public static void sendEmail(String email) {
		
		System.out.print(email + " ....... ");
		try {
			Session session = null;
	        Properties prop = new Properties();
	        prop.put("mail.smtp.auth", true);
	        prop.put("mail.smtp.starttls.enable", "true");
	        prop.put("mail.smtp.host", getHostMail());
	        prop.put("mail.smtp.port", getPortMail());
	       // prop.put("mail.smtp.ssl.trust", smtpHost);
	
	        session = Session.getInstance(prop, new Authenticator() {
	            protected PasswordAuthentication getPasswordAuthentication() {
	                return new PasswordAuthentication(getUsernameMail(), getPasswordMail());
	            }
	        });
	        
	        Message message = new MimeMessage(session);
	        message.setFrom(new InternetAddress(getEmailAddressFrom()));   
	        message.setSentDate(new Date(System.currentTimeMillis()));
//	        message.setRecipients(Message.RecipientType.TO, InternetAddress.parse(email));   
	        message.addRecipient(RecipientType.TO, new InternetAddress(email));
	        message.setSubject("test");
	        
	        String msg = "test covid";
	        
	        MimeBodyPart mimeBodyPart = new MimeBodyPart();
	        MimeBodyPart imagePart = new MimeBodyPart();
	        Multipart multipart = new MimeMultipart();
	        
	        mimeBodyPart.setContent(msg, "text/html");            
	        mimeBodyPart.setHeader("Content-Transfer-Encoding","base64");
	        
	        multipart.addBodyPart(mimeBodyPart);          
	                   
	        message.setContent(multipart);
	        Transport.send(message);
	        System.out.println("Message Sent!");
        
	    } catch (Exception e) {
	        e.printStackTrace();
	    }
	}
}
