/**
 * miniz.c 嵌入式静态实现
 */
#include "miniz.h"
#include <stdio.h>

void mz_free(void *p) {
    if (p) free(p);
}
