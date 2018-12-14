package wsh.io.mail.app;

import java.io.BufferedOutputStream;
import java.io.BufferedReader;
import java.io.ByteArrayInputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Collections;
import java.util.Date;
import java.util.List;
import java.util.Properties;

import javax.activation.DataSource;
import javax.mail.Address;
import javax.mail.Authenticator;
import javax.mail.Folder;
import javax.mail.Message;
import javax.mail.MessagingException;
import javax.mail.PasswordAuthentication;
import javax.mail.Session;
import javax.mail.Store;
import javax.mail.internet.InternetAddress;
import javax.mail.internet.MimeMessage;

import org.apache.commons.mail.util.MimeMessageParser;
import org.apache.pdfbox.io.RandomAccessBuffer;
import org.apache.pdfbox.pdfparser.PDFParser;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.text.PDFTextStripper;

public class App {
	
	private static final String SAVE_PATH = "PDF_" + new SimpleDateFormat("yyyyMMddHHmmss").format(new Date());
	
	static Store store  = null;
	static Folder inbox = null;
	static Session session = null;
		
	public static void main(String[] args) throws Exception {
		try {
			args = new String[] {"", ""};
			
			mkdir(SAVE_PATH);
			Authenticator authenticator = createAuthenticator(args);
			
			MessageFilter msgFilter = new MessageFilter() {
	
				@Override
				public boolean accept(Message message) {
					try {
						Address[] from = message.getFrom();
						if (from != null && from.length > 0) {
							
							String sFrom = ((InternetAddress)from[0]).getAddress().toLowerCase();
							String receivedDate = new SimpleDateFormat("yyyyMM").format(message.getReceivedDate());
							return "2186545300@qq.com".equals(sFrom) && receivedDate.compareTo("201810") >= 0;
						}
					} catch (MessagingException e) {
						e.printStackTrace();
					}
					return false;
				}
			};
			
			List<Message> msgList = listEmails(HostProfile.TENCENT.getProfile(), authenticator, msgFilter);
			System.out.println("TOTAL EMAILS:" + msgList.size());
			AttachementFilter pdfFilter = new AttachementFilter() {
				@Override
				public boolean accept(DataSource dataSource) {
					String filename = dataSource.getName().toLowerCase();
					return filename.endsWith(".pdf");
				}
				
			};
			
			int processingNum = 0;
			for (Message msg : msgList) {
				++processingNum;
				System.out.printf("PROCESSING %d of %d\n", processingNum, msgList.size());
				List<File> pdfList = saveAttachementAs(msg, pdfFilter);
				renamePdfFiles(pdfList);
			}
		} finally {
			try {if(inbox != null)inbox.close();} catch(Exception e) {}
			try {if(store != null)store.close();} catch(Exception e) {}
		}
	}
	
	static void mkdir(String savePath) {
		File path = new File(savePath);
		if (!path.exists()) {
			path.mkdirs();
		}
	}
	
	static void renamePdfFiles(List<File> pdfList) {
		if (pdfList == null) {
			return;
		}
		
		for (File pdfFile : pdfList) {
			String sampleName = getSampleName(pdfFile);
			if (sampleName != null) {
				File dest = determineDestinationFile(pdfFile.getParent(), sampleName);
				pdfFile.renameTo(dest);
			}
		}
	}
	
	static List<File> saveAttachementAs(Message message) {
		return saveAttachementAs(message, null);
	}
	
	static List<File> saveAttachementAs(Message message, AttachementFilter filter) {
		List<File> fileList = new ArrayList<File>();
		try {
			MimeMessageParser parser = new MimeMessageParser((MimeMessage) message).parse();
			List<DataSource> attachedList = parser.getAttachmentList();
			if (attachedList == null)
				return fileList;
			
			if (filter == null) {
				filter = new AttachementFilter() {
					@Override
					public boolean accept(DataSource dataSource) {
						return true;
					}
				};
			}
			
			String receivedDate = new SimpleDateFormat("yyyyMM").format(message.getReceivedDate());
			String savePath = SAVE_PATH + File.separator + receivedDate;
			mkdir(savePath);
			for (DataSource ds : attachedList) {
				if (!filter.accept(ds)) {
					continue;
				}
				String fileName = savePath + File.separator + ds.getName();
				try (BufferedOutputStream os = new BufferedOutputStream(
						new FileOutputStream(fileName))) {
					try (InputStream inputStream = ds.getInputStream()) {
						byte[] buffer = new byte[1024*8];
						int c = -1;
						while ((c = inputStream.read(buffer)) != -1) {
							os.write(buffer, 0, c);
						}
						os.flush();
					} catch (Exception e) {
						e.printStackTrace();
					}
				} catch (Exception e) {
					e.printStackTrace();
				}
				fileList.add(new File(fileName));
			}
		} catch (Exception e) {
			e.printStackTrace();
		}
		return fileList;
	}
	
	static interface AttachementFilter {
		boolean accept(DataSource dataSource);
	}
	
	static Authenticator createAuthenticator(String[] args) {
		if (args == null || args.length < 2) {
			return null;
		}
		
		final String account  = args[0];
		final String password = args[1];
		return new Authenticator() {
			@Override
			public PasswordAuthentication getPasswordAuthentication() {
				return new PasswordAuthentication(account, password);
			}
		};
	}
	
	static List<Message> listEmails(Properties hostProfile, Authenticator authenticator) {
		return listEmails(hostProfile, authenticator, null);
	}
	
	static List<Message> listEmails(Properties hostProfile, Authenticator authenticator, 
			MessageFilter messageFilter) {
		final List<Message> emailList = Collections.synchronizedList(new ArrayList<Message>());
		Message[] messages;
		session = Session.getDefaultInstance(hostProfile, authenticator);
		
		try {
			if (messageFilter == null) {
				messageFilter = new MessageFilter() {
					public boolean accept(Message message) {
						return true;
					}
				};
			}
			store = session.getStore("imap");
			store.connect();
			inbox = store.getFolder("INBOX");
			inbox.open(Folder.READ_ONLY);
			messages = inbox.getMessages(2800, inbox.getMessageCount());
			int count = messages.length;
			System.out.println("MESSAGES COUNT:" + count);
			if (messages != null && messages.length > 0) {
				for (final Message message : messages) {
					if (messageFilter.accept(message)) {
						emailList.add(message);
					}
					--count;
					System.out.println("LEFT COUNT:" + count);
				}
			}
		} catch (Exception e) {
			e.printStackTrace();
		} finally {
			//try {if (store != null) store.close();} catch(Exception e) {}
		}
		
		return emailList;
	}
	
	static interface MessageFilter {
		boolean accept(Message message);
	}
	
	static enum HostProfile {
		
		TENCENT {
			public Properties getProfile() {
				Properties profile = new Properties();
				profile.put("mail.pop3.host", "pop.exmail.qq.com");
				profile.put("mail.imap.host", "imap.exmail.qq.com");
				profile.put("mail.store.protocol", "pop3");
				return profile;
			}
		};
		
		public abstract Properties getProfile();
	}
	
	static File determineDestinationFile(String savePath, String sampleName) {
		File file = new File(savePath + File.separator + sampleName + ".pdf");
		if (!file.exists()) {
			return file;
		}
		int count = 1;
		for (;;) {
			file = new File(sampleName + "(" + count + ").pdf");
			if (!file.exists()) {
				return file;
			}
			++count;
		}
	}
	
	static String getSampleName(String pdfFileName) {
		if (pdfFileName != null && !"".equals(pdfFileName.trim())) {
			return getSampleName(pdfFileName);
		}
		return null;
	}
	
	static String getSampleName(File pdfFile) {
		if (pdfFile.isFile()) {
			InputStream inputStream = null;
			try {
				inputStream = new FileInputStream(pdfFile);
				return getSampleName(inputStream);
			} catch (Exception e) {
				e.printStackTrace();
			} finally {
				try {
					inputStream.close();
				} catch (Exception e) {
				}
			}
		}
		return null;
	}
	
	static String getSampleName(InputStream inputStream) {
		PDDocument doc = null;
		
		try {
			PDFParser parser = new PDFParser(new RandomAccessBuffer(inputStream));
			parser.parse();
			doc = parser.getPDDocument();
			PDFTextStripper txtStriper = new PDFTextStripper();
			txtStriper.setSortByPosition(true);
			txtStriper.setStartPage(1);
			txtStriper.setEndPage(1);
			
			String content = txtStriper.getText(doc);
			//System.out.println(content);
			
			BufferedReader reader = new BufferedReader(
					new InputStreamReader(new ByteArrayInputStream(content.getBytes("utf-8"))));
			String line = null;
			while ((line = reader.readLine()) != null) {
				if (line.indexOf("样 品 名 称") != -1) {
					String[] els = line.split("：");
					if (els.length > 1) {
						return els[1].trim();
					}
				}
			}
		} catch (Exception e) {
			e.printStackTrace();
		} finally {
			try {doc.close();} catch (Exception e) {}
		}
		
		return null;
	}

}
