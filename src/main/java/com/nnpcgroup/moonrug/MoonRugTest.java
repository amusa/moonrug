/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package com.nnpcgroup.moonrug;

import java.util.HashMap;
import java.util.Map;
import java.util.logging.Level;
import java.util.logging.Logger;

/**
 *
 * @author 18359
 */
public class MoonRugTest {
    // Session session = IEws Session. Factory .create (
    /**
     * @param args the command line arguments
     */
    public static void main(String[] args) {
        // TODO code application logic here
        new MoonRugTest().retrieveEmail();
    }

    public void retrieveEmail() {
        try {
            // JavaMail API - Exchange server to allow IMAP access

            // mail server connection parameters
            String user = "18359";
            String password = "foxjuser";
            String host = "outlook.nnpcgroup.com";
            String domain = "chq";

            Map<String, String> map = new HashMap<>();
            map.put(Session.USERNAME, user);
            map.put(Session.PASSWORD, password);
            map.put(Session.DOMAIN, domain);
            map.put(Session.SERVER, host);
            Session session = IMapiSession.Factory.create(map);

            IFolder inbox = null;

            inbox = session.getStore().getFolder(DefaultFolder.INBOX);
            int totalMsg = inbox.getContentCount();
            int unread = inbox.getUnreadCount();

            System.out.println("Total msg: " + totalMsg + " Unread: " + unread);
        } catch (ExchangeException ex) {
            Logger.getLogger(MoonrugTest.class.getName()).log(Level.SEVERE, null, ex);
        }
    }
}
