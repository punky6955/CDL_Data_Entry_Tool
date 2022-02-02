package org.jis;

import javax.swing.*;
import javax.swing.border.Border;
import java.awt.*;
import java.awt.event.MouseAdapter;
import java.awt.event.MouseEvent;
import java.awt.event.MouseMotionAdapter;

public class cdlForm extends JFrame{
    //create jrame,panel,labels, border, and mouselistener variables
    static JFrame frame = new JFrame();
    JPanel panel = new JPanel(new BorderLayout());
    static JLabel label1 = new JLabel();
    static JLabel label2 = new JLabel();
    static JLabel label3 = new JLabel();
    private int frameX,frameY;
    Border blackline = BorderFactory.createLineBorder(Color.black,2);

    public cdlForm(){
        initialize();
    }

    public void initialize(){
        //add panel to frame,
        frame.add(panel);
        frame.setTitle("CDL/Nespresso");
        frame.setSize(250,150);
        frame.setLocationRelativeTo(null);
        frame.setResizable(false);
        frame.setDefaultCloseOperation(EXIT_ON_CLOSE);
        //remove default border and max,min,close buttons
        frame.setUndecorated(true);
        //add mouselistener so user can move GUI
        frame.addMouseListener(new MouseAdapter() {
            @Override
            public void mousePressed(MouseEvent e) {
                super.mousePressed(e);
                frameX = e.getXOnScreen() - frame.getX();
                frameY = e.getYOnScreen() - frame.getY();
            }
        });
        frame.addMouseMotionListener(new MouseMotionAdapter() {
            @Override
            public void mouseDragged(MouseEvent e) {
                super.mouseDragged(e);
                frame.setLocation(e.getXOnScreen()- frameX, e.getYOnScreen() - frameX);
            }
        });

        //add labels and set location
        panel.add(label1, BorderLayout.PAGE_START);
        panel.add(label2,BorderLayout.CENTER);
        panel.add(label3, BorderLayout.PAGE_END);
        panel.setBackground(new Color(247,244,250));
        panel.setBorder(blackline);

        //set label alignment and look
        label1.setHorizontalAlignment(JLabel.CENTER);
        label2.setHorizontalAlignment(JLabel.CENTER);
        label3.setHorizontalAlignment(JLabel.CENTER);
        label1.setFont(new Font("Rockwell",Font.PLAIN,24));
        label2.setFont(new Font("Rockwell",Font.PLAIN,24));
        label3.setFont(new Font("Rockwell",Font.PLAIN,18));
        label1.setForeground(new Color(31,7,60));
        label2.setForeground(new Color(31,7,60));
        label3.setForeground(new Color(31,7,60));
    }

    //setters for labels
    public static void setTxt1(String txt){
        label1.setText(txt);
    }

    public static void setTxt2(String txt){
        label2.setText(txt);
    }

    public static void setTxt3(String txt){
        label3.setText(txt);
    }

    public static void run(){
        frame.setVisible(true);
    }
}
