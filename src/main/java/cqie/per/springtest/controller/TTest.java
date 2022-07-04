package cqie.per.springtest.controller;

import java.util.Arrays;

public class TTest {
    public static void main(String[] args) {
        for(int i = 0 ;i <= 10 ; i++){
            System.out.println(i);
        }
        int[] array = {1,2,3,4,5,6,7,8,9,10};

        for (int y : array){
            System.out.println(y);
        }

        Arrays.stream(array).forEach(x->System.out.println(x));
    }
}
