package com.dabenxiang.properties;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.boot.context.properties.ConfigurationProperties;
import org.springframework.stereotype.Component;

/**
 * Date:2018/8/25
 * Author: yc.guo the one whom in nengxun
 * Desc:
 */
@Component
@ConfigurationProperties(prefix = "dabenxiang.doc")
public class DocProperties {

    public DocTitleProperties docTitleProperties;

    public DocTitleProperties getDocTitleProperties() {
        return docTitleProperties;
    }

    public void setDocTitleProperties(DocTitleProperties docTitleProperties) {
        this.docTitleProperties = docTitleProperties;
    }
}
