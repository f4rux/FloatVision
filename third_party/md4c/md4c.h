#ifndef MD4C_H
#define MD4C_H

#include <stddef.h>

#ifdef __cplusplus
extern "C" {
#endif

typedef char MD_CHAR;
typedef size_t MD_SIZE;

typedef void (*MD_OUTPUT_FUNC)(const MD_CHAR* text, MD_SIZE size, void* userdata);

int md_parse(const MD_CHAR* text, MD_SIZE size, MD_OUTPUT_FUNC output, void* userdata, unsigned parser_flags);

#ifdef __cplusplus
}
#endif

#endif
