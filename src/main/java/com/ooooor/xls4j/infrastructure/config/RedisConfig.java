package com.ooooor.xls4j.infrastructure.config;

import com.ooooor.xls4j.application.dto.OneLineResultDto;
import org.springframework.beans.factory.annotation.Qualifier;
import org.springframework.context.annotation.Bean;
import org.springframework.context.annotation.Configuration;
import org.springframework.data.redis.connection.lettuce.LettuceConnectionFactory;
import org.springframework.data.redis.core.RedisTemplate;

import java.util.Map;

/**
 * @description:
 * @author: chenr
 * @date: 18-11-11
 */
@Configuration
public class RedisConfig {

    @Bean
    public LettuceConnectionFactory lettuceConnectionFactory() {
        return new LettuceConnectionFactory();
    }

    @Bean
    @Qualifier("LineResultRedisTemplate")
    RedisTemplate< String, OneLineResultDto> LineResultRedisTemplate() {
        final RedisTemplate< String, OneLineResultDto > template = new RedisTemplate<>();
        template.setConnectionFactory( lettuceConnectionFactory() );
        return template;
    }
    @Bean
    @Qualifier("LineMapRedisTemplate")
    RedisTemplate< String, Map<String, String[]>> LineMaptRedisTemplate() {
        final RedisTemplate< String, Map<String, String[]> > template = new RedisTemplate<>();
        template.setConnectionFactory( lettuceConnectionFactory() );
        return template;
    }


}
