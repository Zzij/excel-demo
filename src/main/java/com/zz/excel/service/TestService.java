package com.zz.excel.service;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.scheduling.annotation.Scheduled;
import org.springframework.stereotype.Service;

@Service
public class TestService {

    private Logger logger = LoggerFactory.getLogger(TestService.class);

    @Scheduled(cron = "0/5 * * * * ?")
    public void loggerTest(){
        for(int i = 0; i < 20; i++){
            logger.info("我们的距离好像忽远又忽近，是不是有一种特殊感情");
        }
    }
}
