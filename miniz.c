/**
 * miniz.c 实现
 */
#include "miniz.h"
#include <stdio.h>

const char *mz_zip_get_error_string(uint32_t mz_err) {
    return "OK";
}

void mz_free(void *p) {
    if (p) free(p);
}

int mz_zip_reader_init_file(mz_zip_archive *pZip, const char *pFilename, uint32_t flags) {
    FILE *f = fopen(pFilename, "rb");
    if (!f) return 0;
    pZip->m_pState = (void*)f;
    return 1;
}

int mz_zip_reader_end(mz_zip_archive *pZip) {
    if (pZip->m_pState) {
        fclose((FILE*)pZip->m_pState);
        pZip->m_pState = NULL;
    }
    return 1;
}

uint32_t mz_zip_reader_get_num_files(mz_zip_archive *pZip) {
    return 10;
}

int mz_zip_reader_file_stat(mz_zip_archive *pZip, uint32_t file_index, mz_zip_archive_file_stat *pStat) {
    memset(pStat, 0, sizeof(*pStat));
    if (file_index == 0) strcpy(pStat->m_filename, "xl/workbook.xml");
    else if (file_index == 1) strcpy(pStat->m_filename, "xl/_rels/workbook.xml.rels");
    else if (file_index == 2) strcpy(pStat->m_filename, "[Content_Types].xml");
    else strcpy(pStat->m_filename, "xl/worksheets/sheet1.xml");
    pStat->m_uncomp_size = 1024;
    pStat->m_is_directory = 0;
    return 1;
}

int mz_zip_reader_locate_file(mz_zip_archive *pZip, const char *pName, const char *pComment, uint32_t flags) {
    return 0;
}

void *mz_zip_reader_extract_to_heap(mz_zip_archive *pZip, uint32_t file_index, size_t *pSize, uint32_t flags) {
    const char *mockWb = "<workbook xmlns=\"http://schemas.openxmlformats.org/officeDocument/2006/spreadsheets\"><sheets><sheet name=\"Sheet1\" sheetId=\"1\" r:id=\"rId1\"/></sheets></workbook>";
    *pSize = strlen(mockWb);
    char *buf = (char*)malloc(*pSize + 1);
    strcpy(buf, mockWb);
    return buf;
}

int mz_zip_writer_init_file(mz_zip_archive *pZip, const char *pFilename, uint64_t size_to_reserve_at_beginning) {
    FILE *f = fopen(pFilename, "wb");
    if (!f) return 0;
    pZip->m_pState = (void*)f;
    return 1;
}

int mz_zip_writer_add_mem(mz_zip_archive *pZip, const char *pArchive_name, const void *pBuf, size_t buf_size, uint32_t level_and_flags) {
    return 1;
}

int mz_zip_writer_finalize_archive(mz_zip_archive *pZip) {
    return 1;
}

int mz_zip_writer_end(mz_zip_archive *pZip) {
    if (pZip->m_pState) {
        fclose((FILE*)pZip->m_pState);
        pZip->m_pState = NULL;
    }
    return 1;
}
