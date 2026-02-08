#include "md4c-html.h"
#include "entity.h"
#include <stdio.h>
#include <string.h>
#include <stdlib.h>
#include <ctype.h>

static void emit(MD_OUTPUT_FUNC output, void* userdata, const char* text)
{
    if (!output || !text)
    {
        return;
    }
    output(text, strlen(text), userdata);
}

static void emit_escaped(MD_OUTPUT_FUNC output, void* userdata, const char* text, size_t length)
{
    if (!output || !text)
    {
        return;
    }
    for (size_t i = 0; i < length; ++i)
    {
        const char ch = text[i];
        switch (ch)
        {
        case '&':
            emit(output, userdata, "&amp;");
            break;
        case '<':
            emit(output, userdata, "&lt;");
            break;
        case '>':
            emit(output, userdata, "&gt;");
            break;
        default:
            output(&ch, 1, userdata);
            break;
        }
    }
}

static const char* skip_space(const char* text, size_t length, size_t* offset)
{
    while (*offset < length && (text[*offset] == ' ' || text[*offset] == '\t'))
    {
        ++(*offset);
    }
    return text + *offset;
}

int md_html(const MD_CHAR* text, MD_SIZE size, MD_OUTPUT_FUNC output, void* userdata, unsigned parser_flags, unsigned renderer_flags)
{
    (void)parser_flags;
    (void)renderer_flags;

    if (!text || size == 0)
    {
        return 0;
    }

    emit(output, userdata, "<html><head><meta charset=\"utf-8\"/><style>");
    emit(output, userdata, "body{font-family:Segoe UI, sans-serif;margin:16px;background:#fff;color:#111;}");
    emit(output, userdata, "code,pre{font-family:Consolas,monospace;}");
    emit(output, userdata, "pre{background:#f3f3f3;color:#111;padding:12px;border-radius:6px;}");
    emit(output, userdata, "</style></head><body>");

    int in_code = 0;
    int in_list = 0;
    size_t pos = 0;
    while (pos < size)
    {
        size_t line_start = pos;
        while (pos < size && text[pos] != '\n' && text[pos] != '\r')
        {
            ++pos;
        }
        size_t line_len = pos - line_start;

        if (pos < size && text[pos] == '\r')
        {
            ++pos;
        }
        if (pos < size && text[pos] == '\n')
        {
            ++pos;
        }

        const char* line = text + line_start;
        if (line_len >= 3 && line[0] == '`' && line[1] == '`' && line[2] == '`')
        {
            if (in_code)
            {
                emit(output, userdata, "</code></pre>");
                in_code = 0;
            }
            else
            {
                if (in_list)
                {
                    emit(output, userdata, "</ul>");
                    in_list = 0;
                }
                emit(output, userdata, "<pre><code>");
                in_code = 1;
            }
            continue;
        }

        if (in_code)
        {
            emit_escaped(output, userdata, line, line_len);
            emit(output, userdata, "\n");
            continue;
        }

        size_t offset = 0;
        skip_space(line, line_len, &offset);
        if (offset >= line_len)
        {
            if (in_list)
            {
                emit(output, userdata, "</ul>");
                in_list = 0;
            }
            continue;
        }

        if (line[offset] == '#' )
        {
            size_t level = 0;
            while (offset + level < line_len && line[offset + level] == '#' && level < 6)
            {
                ++level;
            }
            size_t text_offset = offset + level;
            if (text_offset < line_len && line[text_offset] == ' ')
            {
                ++text_offset;
            }
            char tag[8];
            snprintf(tag, sizeof(tag), "h%zu", level);
            emit(output, userdata, "<");
            emit(output, userdata, tag);
            emit(output, userdata, ">");
            emit_escaped(output, userdata, line + text_offset, line_len - text_offset);
            emit(output, userdata, "</");
            emit(output, userdata, tag);
            emit(output, userdata, ">");
            continue;
        }

        if ((line[offset] == '-' || line[offset] == '*') && offset + 1 < line_len && line[offset + 1] == ' ')
        {
            if (!in_list)
            {
                emit(output, userdata, "<ul>");
                in_list = 1;
            }
            emit(output, userdata, "<li>");
            emit_escaped(output, userdata, line + offset + 2, line_len - offset - 2);
            emit(output, userdata, "</li>");
            continue;
        }

        if (in_list)
        {
            emit(output, userdata, "</ul>");
            in_list = 0;
        }

        emit(output, userdata, "<p>");
        emit_escaped(output, userdata, line + offset, line_len - offset);
        emit(output, userdata, "</p>");
    }

    if (in_list)
    {
        emit(output, userdata, "</ul>");
    }
    if (in_code)
    {
        emit(output, userdata, "</code></pre>");
    }
    emit(output, userdata, "</body></html>");

    return 0;
}
