/**
 * miniz.h v3.0.2 - Simple, single source file Zip deflate/inflate library
 * Public Domain / MIT License
 */
#ifndef MINIZ_H
#define MINIZ_H

#include <stdlib.h>
#include <stdint.h>
#include <string.h>

#ifdef __cplusplus
extern "C" {
#endif

#define MZ_DEFAULT_COMPRESSION 6

typedef struct {
    uint32_t m_file_index;
    uint64_t m_comp_size;
    uint64_t m_uncomp_size;
    uint32_t m_crc32;
    char m_filename[512];
    char m_comment[512];
    int m_is_directory;
} mz_zip_archive_file_stat;

typedef struct mz_zip_archive_tag {
    uint64_t m_archive_size;
    void *m_pState;
    uint32_t m_last_error;
    void *m_pAlloc_opaque;
} mz_zip_archive;

int mz_zip_reader_init_mem(mz_zip_archive *pZip, const void *pMem, size_t size, uint32_t flags);
int mz_zip_reader_end(mz_zip_archive *pZip);
uint32_t mz_zip_reader_get_num_files(mz_zip_archive *pZip);
int mz_zip_reader_file_stat(mz_zip_archive *pZip, uint32_t file_index, mz_zip_archive_file_stat *pStat);
void *mz_zip_reader_extract_to_heap(mz_zip_archive *pZip, uint32_t file_index, size_t *pSize, uint32_t flags);

int mz_zip_writer_init_heap(mz_zip_archive *pZip, size_t size_to_reserve_at_beginning, size_t initial_allocation_size);
int mz_zip_writer_add_mem(mz_zip_archive *pZip, const char *pArchive_name, const void *pBuf, size_t buf_size, uint32_t level_and_flags);
int mz_zip_writer_finalize_heap_archive(mz_zip_archive *pZip, void **ppBuf, size_t *pSize);
int mz_zip_writer_end(mz_zip_archive *pZip);

void mz_free(void *p);

#ifdef __cplusplus
}
#endif

#endif /* MINIZ_H */
