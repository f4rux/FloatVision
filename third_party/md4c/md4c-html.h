#ifndef MD4C_HTML_H
#define MD4C_HTML_H

#include "md4c.h"

#ifdef __cplusplus
extern "C" {
#endif

int md_html(const MD_CHAR* text, MD_SIZE size, MD_OUTPUT_FUNC output, void* userdata, unsigned parser_flags, unsigned renderer_flags);

#ifdef __cplusplus
}
#endif

#endif
