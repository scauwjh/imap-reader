package me.keiwu.common.utils;

import com.google.common.collect.Lists;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import javax.mail.*;
import javax.mail.internet.MimeMessage;
import javax.mail.internet.MimeMultipart;
import javax.mail.internet.MimeUtility;
import javax.mail.search.FlagTerm;
import java.io.IOException;
import java.io.UnsupportedEncodingException;
import java.util.List;
import java.util.Properties;

/**
 * Created by kei on 9/22/16.
 */
public class IMAPEmailReader {

    private static final Logger logger = LoggerFactory.getLogger(IMAPEmailReader.class);

    public static final String DEFAULT_MAILBOX = "INBOX";
    public static final String TRANSFER_ENCODING_DEFAULT = "quoted-printable";
    public static final String TRANSFER_ENCODING_BASE64 = "Base64";

    private MimeMessage[] readingMessages;
    private Properties properties;
    private Session session;
    private Store store;
    private Folder folder;

    private boolean isInit;
    private String transferEncoding;
    private String email;
    private String password;

    // visitable as getter and setter
    private StringBuilder textContent;      // 文本内容
    private StringBuilder htmlContent;      // 超文本内容
    private List<BodyPart> sources;         // 内嵌资源
    private List<BodyPart> attachments;     // 附件


    public IMAPEmailReader(String host, int port, String email, String password)  {
        Properties props = new Properties();
        String protocol = host.substring(0, host.indexOf("."));
        props.put("mail.store.protocol", protocol);
        props.put("mail.imap.host", host);
        props.put("mail.imap.port", port);
        props.put("mail.imap.partialfetch", false);
        this.properties = props;
        this.email = email;
        this.password = password;
        this.isInit = false;
        this.transferEncoding = TRANSFER_ENCODING_DEFAULT;
    }


    /**
     * 初始化,打开session连接邮箱
     * @throws MessagingException
     */
    public void init() throws MessagingException {
        this.session = Session.getDefaultInstance(this.properties);
        try {
            this.store = this.session.getStore();
            this.store.connect(this.email, this.password);
            isInit = true;
        } catch (MessagingException e) {
            e.printStackTrace();
            logger.error("IMAPReader init error, email={}, password={}", email, password);
            throw e;
        }
    }

    /**
     * 当 header:Content-Transfer-Encoding 值为空时, 解析时会设置上去
     * <br/>
     * <p>默认: "Quoted-printable"</p>
     * <p>可选: "Base64"</p>
     * @param encoding
     */
    public void setTransferEncoding(String encoding) {
        this.transferEncoding = encoding;
    }

    /**
     * 读取未读邮件,并设置为已读
     * @return
     * @throws MessagingException
     */
    public List<MimeMessage> getUnseenEmailAndMarkSeen() throws MessagingException {
        checkInit();
        if (this.folder == null || !this.folder.isOpen()) {
            useDefaultFolder(Folder.READ_WRITE);
        }
        List<MimeMessage> list = Lists.newArrayList();
        this.readingMessages = (MimeMessage[]) this.folder.search(new FlagTerm(new Flags(Flags.Flag.SEEN), false));
        for (MimeMessage msg : this.readingMessages) {
            markSeenFlag(msg, true);
            MimeMessage newMsg = new MimeMessage(msg);
            list.add(newMsg);
        }
        return list;
    }

    /**
     * 使用文件夹(邮件箱)
     * @param mode
     * @throws MessagingException
     */
    public void useDefaultFolder(int mode) throws MessagingException {
        checkInit();
        useFolder(DEFAULT_MAILBOX, mode);
    }

    /**
     * 使用文件夹(邮件箱)
     * @param mailbox
     * @param mode
     * @throws MessagingException
     */
    public void useFolder(String mailbox,int mode) throws MessagingException {
        checkInit();
        this.folder = store.getFolder(mailbox);
        this.folder.open(mode);
    }

    /**
     * 通过读取中的邮件的顺序index设置读取flag
     * @param index
     * @param flag
     * @return
     */
    public boolean markSeenFlagByReadingMessageIndex(int index, boolean flag) {
        if (index < readingMessages.length) {
            markSeenFlag(readingMessages[index], flag);
            return true;
        }
        return false;
    }

    /**
     * 通过邮件number设置读取flag
     * @param messageNum
     * @param flag
     * @return
     */
    public boolean markSeenFlagByMessageNum(int messageNum, boolean flag) {
        for (Message msg : readingMessages) {
            if (msg.getMessageNumber() == messageNum) {
                markSeenFlag(msg, flag);
                return true;
            }
        }
        return false;
    }

    /**
     * 设置邮件读取flag
     * @param msg
     * @param flag
     * @throws MessagingException
     */
    public void markSeenFlag(Message msg, boolean flag) {
        try {
            msg.setFlag(Flags.Flag.SEEN, flag);
        } catch (MessagingException e) {
            e.printStackTrace();
            logger.error("mark email seen flag error", e);
        }
    }

    /**
     * 获取标题
     * @param msg
     * @return
     * @throws MessagingException
     * @throws UnsupportedEncodingException
     */
    public String getSubject(Message msg) throws MessagingException, UnsupportedEncodingException {
        return MimeUtility.decodeText(msg.getSubject());
    }

    /**
     * 获取邮件内容, 通过其他方法获取实际内容
     * @param msg
     * @throws MessagingException
     * @throws IOException
     */
    public void getContent(Message msg) throws MessagingException, IOException {
        // 初始化参数
        this.initContentParams();
        Object content = msg.getContent();
        if (content instanceof MimeMultipart) {
            MimeMultipart mp = (MimeMultipart) content;
            int count = mp.getCount();
            for (int i = 0; i < count; i++) {
                BodyPart part = mp.getBodyPart(i);
                getOuterParts(part);
            }
        } else if (content instanceof BodyPart) {
            BodyPart part = (BodyPart) content;
            if (part.isMimeType("text/html")) {
                addContentTransferEncoding(part);
                this.htmlContent.append(part.getContent().toString());
            } else if (part.isMimeType("text/*")) {
                addContentTransferEncoding(part);
                this.textContent.append(part.getContent().toString());
            }
        } else if (content instanceof String) {
            this.htmlContent.append(content);
        }
    }

    /**
     * 关闭
     */
    public void close() {
        if (this.folder != null && this.folder.isOpen()) {
            try {
                this.folder.close(false);
            } catch (MessagingException e) {
                logger.warn("Exception when close the folder");
            }
        }
        if (this.store != null && this.store.isConnected()) {
            try {
                this.store.close();
            } catch (MessagingException e) {
                logger.warn("Exception when close the store");
            }
        }
        this.session = null;
        this.properties = null;
        this.readingMessages = null;
    }


    /** getter setter **/
    public StringBuilder getTextContent() {
        return textContent;
    }

    public void setTextContent(StringBuilder textContent) {
        this.textContent = textContent;
    }

    public StringBuilder getHtmlContent() {
        return htmlContent;
    }

    public void setHtmlContent(StringBuilder htmlContent) {
        this.htmlContent = htmlContent;
    }

    public List<BodyPart> getSources() {
        return sources;
    }

    public void setSources(List<BodyPart> sources) {
        this.sources = sources;
    }

    public List<BodyPart> getAttachments() {
        return attachments;
    }

    public void setAttachments(List<BodyPart> attachments) {
        this.attachments = attachments;
    }




    /** private methods **/
    private void initContentParams() {
        if (textContent == null) textContent = new StringBuilder();
        if (htmlContent == null) htmlContent = new StringBuilder();
        if (sources == null) sources = Lists.newArrayList();
        if (attachments == null) attachments = Lists.newArrayList();
        textContent.setLength(0);
        htmlContent.setLength(0);
        sources.clear();
        attachments.clear();
    }

    private void getOuterParts(BodyPart bodyPart)
            throws MessagingException, IOException {
        if (bodyPart.isMimeType("text/html")) {
            addContentTransferEncoding(bodyPart);
            this.htmlContent.append(bodyPart.getContent().toString());
        } else if (bodyPart.isMimeType("text/*")) {
            addContentTransferEncoding(bodyPart);
            this.textContent.append(bodyPart.getContent().toString());
        } else if (bodyPart.isMimeType("multipart/related")) {
            getRelatedParts(bodyPart);
        } else if (bodyPart.isMimeType("multipart/alternative")) {
            getAlternativeParts(bodyPart);
        } else {
            this.attachments.add(bodyPart);
        }
    }

    private void getRelatedParts(BodyPart bodyPart) throws MessagingException, IOException {
        MimeMultipart mp = (MimeMultipart) bodyPart.getContent();
        int count = mp.getCount();
        for (int i = 0; i < count; i++) {
            BodyPart part = mp.getBodyPart(i);
            if (part.isMimeType("multipart/alternative")) {
                getAlternativeParts(part);
            } else {
                this.sources.add(part);
            }
        }
    }

    private void getAlternativeParts(BodyPart bodyPart) throws MessagingException, IOException {
        MimeMultipart mp = (MimeMultipart) bodyPart.getContent();
        int count = mp.getCount();
        for (int i = 0; i < count; i++) {
            BodyPart part = mp.getBodyPart(i);
            if (part.isMimeType("text/html")) {
                addContentTransferEncoding(part);
                this.htmlContent.append(part.getContent().toString());
            } else if (part.isMimeType("text/*")) {
                addContentTransferEncoding(part);
                this.textContent.append(part.getContent().toString());
            }
        }
    }

    private void addContentTransferEncoding(BodyPart bodyPart) throws MessagingException {
        if (bodyPart.getHeader("Content-Transfer-Encoding") == null) {
            bodyPart.setHeader("Content-Transfer-Encoding", this.transferEncoding);
        }
    }

    private void checkInit() throws MessagingException {
        if (this.isInit) return;
        throw new MessagingException("Instance is not init, call init method to init");
    }

}
