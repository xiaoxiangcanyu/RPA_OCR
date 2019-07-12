package ProgressController;

import java.io.IOException;


public class progress {
    public static void main(String args[]) throws IOException {
        for(int i=0;i<50;i++){
            System.out.println("打游戏"+i);
            if (i==10){
               new Thread().start();
            }
        }

    }



}
