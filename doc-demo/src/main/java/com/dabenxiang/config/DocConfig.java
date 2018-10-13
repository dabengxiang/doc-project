package com.dabenxiang.config;

import com.dabenxiang.properties.DocProperties;
import org.springframework.boot.context.properties.EnableConfigurationProperties;
import org.springframework.context.annotation.Configuration;

/**
 * Date:2018/8/25
 * Author: yc.guo the one whom in nengxun
 * Desc:
 */
@Configuration
@EnableConfigurationProperties(DocProperties.class)
public class DocConfig {
}
