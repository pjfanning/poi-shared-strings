package com.github.pjfanning.poi.xssf.streaming;

import org.apache.poi.xssf.model.Comments;

public interface CommentsTableFactory {
    /**
     * @return a new {@link Comments} implementation instance, configured to your requirements
     */
    Comments createCommentsTable();
}
