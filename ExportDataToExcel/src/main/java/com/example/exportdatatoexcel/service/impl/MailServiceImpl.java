package com.example.exportdatatoexcel.service.impl;

import com.example.exportdatatoexcel.service.MailService;
import com.example.exportdatatoexcel.service.ThymeleafService;
import com.example.exportdatatoexcel.utils.CustomerDTO;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.mail.javamail.JavaMailSender;
import org.springframework.mail.javamail.MimeMessageHelper;
import org.springframework.stereotype.Service;

import javax.mail.internet.MimeMessage;
import java.nio.charset.StandardCharsets;
import java.text.SimpleDateFormat;
import java.util.Arrays;
import java.util.Date;
import java.util.HashMap;
import java.util.Map;

@Service
public class MailServiceImpl implements MailService {

    @Autowired
    JavaMailSender mailSender;

    @Autowired
    ThymeleafService thymeleafService;

    @Value("${spring.mail.username}")
    private String email;

    @Override
    public void sendMailTest() {
        try {
            MimeMessage message = mailSender.createMimeMessage();

            MimeMessageHelper helper = new MimeMessageHelper(
                    message,
                    MimeMessageHelper.MULTIPART_MODE_MIXED_RELATED,
                    StandardCharsets.UTF_8.name());

            helper.setFrom(email);
            helper.setText(thymeleafService.createContent("mail-sender-test.html", null), true);
            helper.setCc("springbootselfcode@gmail.com");
            helper.setTo("kayteecmc@gmail.com");
            helper.setSubject("Mail Test With Template HTML");

            mailSender.send(message);
        }catch (Exception e){
            e.printStackTrace();
        }
    }

    @Override
    public void sendMailCreateCustomer(CustomerDTO dto) {
        try {
            MimeMessage message = mailSender.createMimeMessage();
            MimeMessageHelper helper = new MimeMessageHelper(
                    message, MimeMessageHelper.MULTIPART_MODE_MIXED_RELATED,
                    StandardCharsets.UTF_8.name()
            );

            helper.setTo(dto.getEmail());

            Object[] bccObject = dto.getMailCC().toArray();
            String[] bcc = Arrays.copyOf(bccObject, bccObject.length, String[].class);
            helper.setBcc(bcc);

            Map<String, Object> variables = new HashMap<>();
            variables.put("full_name", dto.getCustomerName());
            variables.put("username", dto.getUsername());
            variables.put("first_name", dto.getContactFirstName());
            variables.put("last_name", dto.getContactLastName());
            variables.put("email", dto.getEmail());
            SimpleDateFormat sdf = new SimpleDateFormat("dd/MM/yyyy");
            variables.put("date", sdf.format(new Date()));
            helper.setText(thymeleafService.createContent("create-customer-mail-template.html", variables), true);
            helper.setFrom(email);
            mailSender.send(message);
        }catch (Exception e){
            e.printStackTrace();
        }
    }
}
