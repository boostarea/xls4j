package com.ooooor.xls4j.infrastructure.config;

import com.ooooor.xls4j.application.dto.OneLineResultDto;
import org.springframework.beans.factory.annotation.Qualifier;
import org.springframework.context.annotation.Bean;
import org.springframework.context.annotation.Configuration;
import org.springframework.data.redis.connection.lettuce.LettuceConnectionFactory;
import org.springframework.data.redis.core.RedisTemplate;
import org.springframework.data.redis.serializer.GenericToStringSerializer;
import org.springframework.data.redis.serializer.StringRedisSerializer;

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
        final RedisTemplate< String, OneLineResultDto > template =  new RedisTemplate< String, OneLineResultDto >();
        template.setConnectionFactory( lettuceConnectionFactory() );
        template.setKeySerializer( new StringRedisSerializer() );
        template.setHashValueSerializer( new GenericToStringSerializer< OneLineResultDto >( OneLineResultDto.class ) );
        template.setValueSerializer( new GenericToStringSerializer< OneLineResultDto >( OneLineResultDto.class ) );
        return template;
    }
    @Bean
    @Qualifier("LineMaptRedisTemplate")
    RedisTemplate< String, Map<String, String[]>> LineMaptRedisTemplate() {
        final RedisTemplate< String, Map<String, String[]> > template =  new RedisTemplate< String, Map<String, String[]> >();
        template.setConnectionFactory( lettuceConnectionFactory() );
        template.setKeySerializer( new StringRedisSerializer() );
        template.setHashValueSerializer( new GenericToStringSerializer< Map >( Map.class ) );
        template.setValueSerializer( new GenericToStringSerializer< Map >( Map.class ) );
        return template;
    }


}
