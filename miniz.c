/**
 * miniz.c - Self-contained ZIP Archive Engine (RFC 1951 Inflate + PKZIP Reader/Writer)
 * Zero external library dependencies. Compiles cleanly on MSVC / GCC / Clang.
 */

#include "miniz.h"
#include <stdio.h>
#include <stdlib.h>
#include <string.h>
#include <time.h>

/* Compute a valid MS-DOS date/time pair for the current local time.
 * ZIP requires: date bits15-9=year-1980 (1-127), bits8-5=month(1-12), bits4-0=day(1-31)
 *               time bits15-11=hour, bits10-5=minute, bits4-0=second/2
 * A value of 0 for month or day is illegal and causes strict readers
 * (including Microsoft Office's OPC parser) to reject the archive. */
static void mz_get_dos_datetime(uint16_t *out_time, uint16_t *out_date) {
    time_t now = time(NULL);
    struct tm tmv;
#if defined(_WIN32)
    localtime_s(&tmv, &now);
#else
    localtime_r(&now, &tmv);
#endif
    int year = tmv.tm_year + 1900;
    if (year < 1980) year = 1980;      /* DOS date cannot represent < 1980 */
    if (year > 2107) year = 2107;      /* DOS date max representable year */

    *out_date = (uint16_t)(((year - 1980) << 9) |
                            ((tmv.tm_mon + 1) << 5) |  /* tm_mon is 0-11 -> DOS wants 1-12 */
                            (tmv.tm_mday));            /* tm_mday is already 1-31 */
    *out_time = (uint16_t)((tmv.tm_hour << 11) |
                            (tmv.tm_min << 5) |
                            (tmv.tm_sec / 2));
}

void mz_free(void *p) {
    if (p) free(p);
}

/* ========================================================================== */
/* CRC32                                                                      */
/* ========================================================================== */
static uint32_t s_crc32_table[256];
static int s_crc32_initialized = 0;

static void mz_init_crc32(void) {
    if (s_crc32_initialized) return;
    for (uint32_t i = 0; i < 256; i++) {
        uint32_t c = i;
        for (int j = 0; j < 8; j++) {
            c = (c & 1) ? (0xEDB88320u ^ (c >> 1)) : (c >> 1);
        }
        s_crc32_table[i] = c;
    }
    s_crc32_initialized = 1;
}

static uint32_t mz_calc_crc32(const uint8_t *buf, size_t len) {
    mz_init_crc32();
    uint32_t crc = 0xFFFFFFFFu;
    for (size_t i = 0; i < len; i++) {
        crc = s_crc32_table[(crc ^ buf[i]) & 0xFF] ^ (crc >> 8);
    }
    return crc ^ 0xFFFFFFFFu;
}

/* ========================================================================== */
/* TINY INFLATE (RFC 1951 Deflate Decompressor)                               */
/* ========================================================================== */
typedef struct {
    const uint8_t *src;
    size_t src_len;
    size_t src_pos;
    uint32_t bit_buf;
    int bit_cnt;

    uint8_t *dest;
    size_t dest_len;
    size_t dest_pos;
} tinfl_ctx;

static uint32_t tinfl_get_bits(tinfl_ctx *ctx, int num_bits) {
    while (ctx->bit_cnt < num_bits) {
        if (ctx->src_pos >= ctx->src_len) return 0;
        ctx->bit_buf |= ((uint32_t)ctx->src[ctx->src_pos++]) << ctx->bit_cnt;
        ctx->bit_cnt += 8;
    }
    uint32_t val = ctx->bit_buf & ((1u << num_bits) - 1);
    ctx->bit_buf >>= num_bits;
    ctx->bit_cnt -= num_bits;
    return val;
}

typedef struct {
    uint16_t counts[16];
    uint16_t symbols[288];
} huff_tree;

static int build_huffman_tree(huff_tree *tree, const uint8_t *lens, int num_symbols) {
    memset(tree->counts, 0, sizeof(tree->counts));
    for (int i = 0; i < num_symbols; i++) {
        if (lens[i] > 15) return 0;
        tree->counts[lens[i]]++;
    }
    tree->counts[0] = 0;

    uint16_t offsets[16];
    offsets[1] = 0;
    for (int i = 1; i < 15; i++) {
        offsets[i + 1] = offsets[i] + tree->counts[i];
    }

    for (int i = 0; i < num_symbols; i++) {
        if (lens[i] != 0) {
            tree->symbols[offsets[lens[i]]++] = (uint16_t)i;
        }
    }
    return 1;
}

static int decode_symbol(tinfl_ctx *ctx, const huff_tree *tree) {
    int code = 0;
    int first = 0;
    int index = 0;

    for (int len = 1; len <= 15; len++) {
        code |= (int)tinfl_get_bits(ctx, 1);
        int count = tree->counts[len];
        if (code - count < first) {
            return tree->symbols[index + (code - first)];
        }
        index += count;
        first += count;
        first <<= 1;
        code <<= 1;
    }
    return -1;
}

static int inflate_block_uncompressed(tinfl_ctx *ctx) {
    ctx->bit_buf = 0;
    ctx->bit_cnt = 0;
    if (ctx->src_pos + 4 > ctx->src_len) return 0;
    uint16_t len = (uint16_t)(ctx->src[ctx->src_pos] | (ctx->src[ctx->src_pos + 1] << 8));
    ctx->src_pos += 4; // skip len and nlen
    if (ctx->src_pos + len > ctx->src_len) return 0;
    if (ctx->dest_pos + len > ctx->dest_len) return 0;
    memcpy(ctx->dest + ctx->dest_pos, ctx->src + ctx->src_pos, len);
    ctx->src_pos += len;
    ctx->dest_pos += len;
    return 1;
}

static const uint16_t k_len_base[29] = {
    3,4,5,6,7,8,9,10,11,13,15,17,19,23,27,31,35,43,51,59,67,83,99,115,131,163,195,227,258
};
static const uint8_t k_len_extra[29] = {
    0,0,0,0,0,0,0,0,1,1,1,1,2,2,2,2,3,3,3,3,4,4,4,4,5,5,5,5,0
};
static const uint16_t k_dist_base[30] = {
    1,2,3,4,5,7,9,13,17,25,33,49,65,97,129,193,257,385,513,769,1025,1537,2049,3073,4097,6145,8193,12289,16385,24577
};
static const uint8_t k_dist_extra[30] = {
    0,0,0,0,1,1,2,2,3,3,4,4,5,5,6,6,7,7,8,8,9,9,10,10,11,11,12,12,13,13
};

static int inflate_huffman(tinfl_ctx *ctx, const huff_tree *ltree, const huff_tree *dtree) {
    for (;;) {
        int sym = decode_symbol(ctx, ltree);
        if (sym < 0) return 0;
        if (sym < 256) {
            if (ctx->dest_pos >= ctx->dest_len) return 0;
            ctx->dest[ctx->dest_pos++] = (uint8_t)sym;
        } else if (sym == 256) {
            return 1; // End of block
        } else if (sym <= 285) {
            int len_idx = sym - 257;
            int length = k_len_base[len_idx] + (int)tinfl_get_bits(ctx, k_len_extra[len_idx]);
            int dist_sym = decode_symbol(ctx, dtree);
            if (dist_sym < 0 || dist_sym >= 30) return 0;
            int dist = k_dist_base[dist_sym] + (int)tinfl_get_bits(ctx, k_dist_extra[dist_sym]);
            if ((size_t)dist > ctx->dest_pos) return 0;
            if (ctx->dest_pos + length > ctx->dest_len) return 0;
            for (int i = 0; i < length; i++) {
                ctx->dest[ctx->dest_pos] = ctx->dest[ctx->dest_pos - dist];
                ctx->dest_pos++;
            }
        } else {
            return 0;
        }
    }
}

static int tinfl_decompress(const uint8_t *src, size_t src_len, uint8_t *dest, size_t dest_len) {
    tinfl_ctx ctx;
    memset(&ctx, 0, sizeof(ctx));
    ctx.src = src;
    ctx.src_len = src_len;
    ctx.dest = dest;
    ctx.dest_len = dest_len;

    int final_block = 0;
    while (!final_block) {
        final_block = (int)tinfl_get_bits(&ctx, 1);
        int block_type = (int)tinfl_get_bits(&ctx, 2);

        if (block_type == 0) { // Uncompressed
            if (!inflate_block_uncompressed(&ctx)) return 0;
        } else if (block_type == 1) { // Fixed Huffman
            uint8_t l_lens[288];
            for (int i = 0; i <= 143; i++) l_lens[i] = 8;
            for (int i = 144; i <= 255; i++) l_lens[i] = 9;
            for (int i = 256; i <= 279; i++) l_lens[i] = 7;
            for (int i = 280; i <= 287; i++) l_lens[i] = 8;
            huff_tree ltree, dtree;
            build_huffman_tree(&ltree, l_lens, 288);
            uint8_t d_lens[32];
            for (int i = 0; i < 32; i++) d_lens[i] = 5;
            build_huffman_tree(&dtree, d_lens, 32);
            if (!inflate_huffman(&ctx, &ltree, &dtree)) return 0;
        } else if (block_type == 2) { // Dynamic Huffman
            int hlit = (int)tinfl_get_bits(&ctx, 5) + 257;
            int hdist = (int)tinfl_get_bits(&ctx, 5) + 1;
            int hclen = (int)tinfl_get_bits(&ctx, 4) + 4;

            static const uint8_t k_cl_order[19] = {
                16, 17, 18, 0, 8, 7, 9, 6, 10, 5, 11, 4, 12, 3, 13, 2, 14, 1, 15
            };
            uint8_t code_lens[19] = { 0 };
            for (int i = 0; i < hclen; i++) {
                code_lens[k_cl_order[i]] = (uint8_t)tinfl_get_bits(&ctx, 3);
            }
            huff_tree ctree;
            build_huffman_tree(&ctree, code_lens, 19);

            uint8_t dynamic_lens[288 + 32] = { 0 };
            int total_syms = hlit + hdist;
            int cur_sym = 0;
            while (cur_sym < total_syms) {
                int sym = decode_symbol(&ctx, &ctree);
                if (sym < 0) return 0;
                if (sym < 16) {
                    dynamic_lens[cur_sym++] = (uint8_t)sym;
                } else if (sym == 16) {
                    if (cur_sym == 0) return 0;
                    uint8_t prev = dynamic_lens[cur_sym - 1];
                    int rep = (int)tinfl_get_bits(&ctx, 2) + 3;
                    while (rep-- && cur_sym < total_syms) dynamic_lens[cur_sym++] = prev;
                } else if (sym == 17) {
                    int rep = (int)tinfl_get_bits(&ctx, 3) + 3;
                    while (rep-- && cur_sym < total_syms) dynamic_lens[cur_sym++] = 0;
                } else if (sym == 18) {
                    int rep = (int)tinfl_get_bits(&ctx, 7) + 11;
                    while (rep-- && cur_sym < total_syms) dynamic_lens[cur_sym++] = 0;
                }
            }

            huff_tree ltree, dtree;
            build_huffman_tree(&ltree, dynamic_lens, hlit);
            build_huffman_tree(&dtree, dynamic_lens + hlit, hdist);
            if (!inflate_huffman(&ctx, &ltree, &dtree)) return 0;
        } else {
            return 0;
        }
    }
    return (ctx.dest_pos == dest_len) ? 1 : 0;
}

/* ========================================================================== */
/* ZIP READER & WRITER ENGINE                                                 */
/* ========================================================================== */

typedef struct {
    const uint8_t *data;
    size_t size;
    uint32_t num_files;
    mz_zip_archive_file_stat *stats;
} zip_reader_state;

int mz_zip_reader_init_mem(mz_zip_archive *pZip, const void *pMem, size_t size, uint32_t flags) {
    (void)flags;
    if (!pZip || !pMem || size < 22) return 0;
    memset(pZip, 0, sizeof(*pZip));
    pZip->m_archive_size = size;

    const uint8_t *buf = (const uint8_t*)pMem;
    size_t eocd_pos = 0;
    int found_eocd = 0;

    size_t scan_limit = (size > 65536 + 22) ? (size - 65536 - 22) : 0;
    for (size_t i = size - 22; i >= scan_limit; i--) {
        if (buf[i] == 0x50 && buf[i + 1] == 0x4b && buf[i + 2] == 0x05 && buf[i + 3] == 0x06) {
            eocd_pos = i;
            found_eocd = 1;
            break;
        }
        if (i == 0) break;
    }
    if (!found_eocd) return 0;

    uint16_t total_entries = (uint16_t)(buf[eocd_pos + 10] | (buf[eocd_pos + 11] << 8));
    uint32_t cd_size = (uint32_t)(buf[eocd_pos + 12] | (buf[eocd_pos + 13] << 8) | (buf[eocd_pos + 14] << 16) | (buf[eocd_pos + 15] << 24));
    uint32_t cd_offset = (uint32_t)(buf[eocd_pos + 16] | (buf[eocd_pos + 17] << 8) | (buf[eocd_pos + 18] << 16) | (buf[eocd_pos + 19] << 24));

    if (cd_offset + cd_size > size) return 0;

    zip_reader_state *state = (zip_reader_state*)calloc(1, sizeof(zip_reader_state));
    if (!state) return 0;
    state->data = buf;
    state->size = size;
    state->num_files = total_entries;
    state->stats = (mz_zip_archive_file_stat*)calloc(total_entries, sizeof(mz_zip_archive_file_stat));

    size_t cur = cd_offset;
    for (uint32_t i = 0; i < total_entries; i++) {
        if (cur + 46 > size) break;
        if (buf[cur] != 0x50 || buf[cur+1] != 0x4b || buf[cur+2] != 0x01 || buf[cur+3] != 0x02) break;

        uint16_t method = (uint16_t)(buf[cur + 10] | (buf[cur + 11] << 8));
        uint32_t crc = (uint32_t)(buf[cur + 16] | (buf[cur + 17] << 8) | (buf[cur + 18] << 16) | (buf[cur + 19] << 24));
        uint32_t comp_sz = (uint32_t)(buf[cur + 20] | (buf[cur + 21] << 8) | (buf[cur + 22] << 16) | (buf[cur + 23] << 24));
        uint32_t uncomp_sz = (uint32_t)(buf[cur + 24] | (buf[cur + 25] << 8) | (buf[cur + 26] << 16) | (buf[cur + 27] << 24));
        uint16_t fn_len = (uint16_t)(buf[cur + 28] | (buf[cur + 29] << 8));
        uint16_t extra_len = (uint16_t)(buf[cur + 30] | (buf[cur + 31] << 8));
        uint16_t comm_len = (uint16_t)(buf[cur + 32] | (buf[cur + 33] << 8));
        uint32_t local_ofs = (uint32_t)(buf[cur + 42] | (buf[cur + 43] << 8) | (buf[cur + 44] << 16) | (buf[cur + 45] << 24));

        state->stats[i].m_file_index = i;
        state->stats[i].m_method = method;
        state->stats[i].m_comp_size = comp_sz;
        state->stats[i].m_uncomp_size = uncomp_sz;
        state->stats[i].m_crc32 = crc;
        state->stats[i].m_local_header_ofs = local_ofs;

        size_t copy_fn = (fn_len < 511) ? fn_len : 511;
        if (cur + 46 + copy_fn <= size) {
            memcpy(state->stats[i].m_filename, buf + cur + 46, copy_fn);
            state->stats[i].m_filename[copy_fn] = '\0';
        }

        if (copy_fn > 0 && state->stats[i].m_filename[copy_fn - 1] == '/') {
            state->stats[i].m_is_directory = 1;
        }

        cur += 46 + fn_len + extra_len + comm_len;
    }

    pZip->m_pState = state;
    return 1;
}

uint32_t mz_zip_reader_get_num_files(mz_zip_archive *pZip) {
    if (!pZip || !pZip->m_pState) return 0;
    return ((zip_reader_state*)pZip->m_pState)->num_files;
}

int mz_zip_reader_file_stat(mz_zip_archive *pZip, uint32_t file_index, mz_zip_archive_file_stat *pStat) {
    if (!pZip || !pZip->m_pState || !pStat) return 0;
    zip_reader_state *state = (zip_reader_state*)pZip->m_pState;
    if (file_index >= state->num_files) return 0;
    *pStat = state->stats[file_index];
    return 1;
}

void *mz_zip_reader_extract_to_heap(mz_zip_archive *pZip, uint32_t file_index, size_t *pSize, uint32_t flags) {
    (void)flags;
    if (!pZip || !pZip->m_pState || !pSize) return NULL;
    zip_reader_state *state = (zip_reader_state*)pZip->m_pState;
    if (file_index >= state->num_files) return NULL;

    const mz_zip_archive_file_stat *stat = &state->stats[file_index];
    uint32_t l_ofs = stat->m_local_header_ofs;
    if (l_ofs + 30 > state->size) return NULL;

    const uint8_t *buf = state->data;
    if (buf[l_ofs] != 0x50 || buf[l_ofs+1] != 0x4b || buf[l_ofs+2] != 0x03 || buf[l_ofs+3] != 0x04) return NULL;

    uint16_t fn_len = (uint16_t)(buf[l_ofs + 26] | (buf[l_ofs + 27] << 8));
    uint16_t extra_len = (uint16_t)(buf[l_ofs + 28] | (buf[l_ofs + 29] << 8));
    size_t data_ofs = l_ofs + 30 + fn_len + extra_len;

    if (data_ofs + stat->m_comp_size > state->size) return NULL;

    uint8_t *out_data = (uint8_t*)malloc((size_t)stat->m_uncomp_size + 1);
    if (!out_data) return NULL;

    if (stat->m_method == 0) { // STORE
        if (stat->m_comp_size != stat->m_uncomp_size) {
            free(out_data);
            return NULL;
        }
        memcpy(out_data, buf + data_ofs, (size_t)stat->m_comp_size);
    } else if (stat->m_method == 8) { // DEFLATE
        if (!tinfl_decompress(buf + data_ofs, (size_t)stat->m_comp_size, out_data, (size_t)stat->m_uncomp_size)) {
            free(out_data);
            return NULL;
        }
    } else {
        free(out_data);
        return NULL;
    }

    out_data[stat->m_uncomp_size] = '\0';
    *pSize = (size_t)stat->m_uncomp_size;
    return out_data;
}

int mz_zip_reader_end(mz_zip_archive *pZip) {
    if (!pZip) return 0;
    if (pZip->m_pState) {
        zip_reader_state *state = (zip_reader_state*)pZip->m_pState;
        if (state->stats) free(state->stats);
        free(state);
        pZip->m_pState = NULL;
    }
    return 1;
}

/* ========================================================================== */
/* ZIP WRITER                                                                 */
/* ========================================================================== */

typedef struct {
    char *name;
    uint8_t *data;
    size_t size;
    uint32_t crc;
    uint32_t offset;
} writer_entry;

typedef struct {
    writer_entry *entries;
    size_t count;
    size_t capacity;
} zip_writer_state;

int mz_zip_writer_init_heap(mz_zip_archive *pZip, size_t size_to_reserve_at_beginning, size_t initial_allocation_size) {
    (void)size_to_reserve_at_beginning;
    (void)initial_allocation_size;
    if (!pZip) return 0;
    memset(pZip, 0, sizeof(*pZip));
    zip_writer_state *state = (zip_writer_state*)calloc(1, sizeof(zip_writer_state));
    if (!state) return 0;
    pZip->m_pState = state;
    return 1;
}

int mz_zip_writer_add_mem(mz_zip_archive *pZip, const char *pArchive_name, const void *pBuf, size_t buf_size, uint32_t level_and_flags) {
    (void)level_and_flags;
    if (!pZip || !pZip->m_pState || !pArchive_name || (!pBuf && buf_size > 0)) return 0;
    zip_writer_state *state = (zip_writer_state*)pZip->m_pState;

    if (state->count >= state->capacity) {
        size_t new_cap = (state->capacity == 0) ? 16 : (state->capacity * 2);
        writer_entry *new_entries = (writer_entry*)realloc(state->entries, new_cap * sizeof(writer_entry));
        if (!new_entries) return 0;
        state->entries = new_entries;
        state->capacity = new_cap;
    }

    writer_entry *e = &state->entries[state->count++];
    size_t nlen = strlen(pArchive_name);
    e->name = (char*)malloc(nlen + 1);
    strcpy(e->name, pArchive_name);

    e->size = buf_size;
    e->data = (uint8_t*)malloc(buf_size + 1);
    if (buf_size > 0 && pBuf) {
        memcpy(e->data, pBuf, buf_size);
        e->crc = mz_calc_crc32((const uint8_t*)pBuf, buf_size);
    } else {
        e->crc = 0;
    }
    return 1;
}

int mz_zip_writer_finalize_heap_archive(mz_zip_archive *pZip, void **ppBuf, size_t *pSize) {
    if (!pZip || !pZip->m_pState || !ppBuf || !pSize) return 0;
    zip_writer_state *state = (zip_writer_state*)pZip->m_pState;

    size_t total_size = 0;
    for (size_t i = 0; i < state->count; i++) {
        size_t nlen = strlen(state->entries[i].name);
        total_size += 30 + nlen + state->entries[i].size;
    }
    size_t cd_offset = total_size;
    size_t cd_size = 0;
    for (size_t i = 0; i < state->count; i++) {
        size_t nlen = strlen(state->entries[i].name);
        cd_size += 46 + nlen;
    }
    total_size += cd_size + 22;

    uint8_t *out = (uint8_t*)malloc(total_size);
    if (!out) return 0;

    uint16_t dos_time, dos_date;
    mz_get_dos_datetime(&dos_time, &dos_date);

    size_t cur = 0;
    for (size_t i = 0; i < state->count; i++) {
        writer_entry *e = &state->entries[i];
        e->offset = (uint32_t)cur;
        size_t nlen = strlen(e->name);

        out[cur + 0] = 0x50; out[cur + 1] = 0x4b; out[cur + 2] = 0x03; out[cur + 3] = 0x04;
        out[cur + 4] = 20; out[cur + 5] = 0;
        out[cur + 6] = 0x00; out[cur + 7] = 0x08; // UTF-8 flag
        out[cur + 8] = 0; out[cur + 9] = 0;       // STORE
        out[cur + 10] = (uint8_t)(dos_time); out[cur + 11] = (uint8_t)(dos_time >> 8);
        out[cur + 12] = (uint8_t)(dos_date); out[cur + 13] = (uint8_t)(dos_date >> 8);
        out[cur + 14] = (uint8_t)(e->crc); out[cur + 15] = (uint8_t)(e->crc >> 8);
        out[cur + 16] = (uint8_t)(e->crc >> 16); out[cur + 17] = (uint8_t)(e->crc >> 24);
        out[cur + 18] = (uint8_t)(e->size); out[cur + 19] = (uint8_t)(e->size >> 8);
        out[cur + 20] = (uint8_t)(e->size >> 16); out[cur + 21] = (uint8_t)(e->size >> 24);
        out[cur + 22] = (uint8_t)(e->size); out[cur + 23] = (uint8_t)(e->size >> 8);
        out[cur + 24] = (uint8_t)(e->size >> 16); out[cur + 25] = (uint8_t)(e->size >> 24);
        out[cur + 26] = (uint8_t)(nlen); out[cur + 27] = (uint8_t)(nlen >> 8);
        out[cur + 28] = 0; out[cur + 29] = 0;
        memcpy(out + cur + 30, e->name, nlen);
        cur += 30 + nlen;
        if (e->size > 0 && e->data) {
            memcpy(out + cur, e->data, e->size);
            cur += e->size;
        }
    }

    size_t cd_cur = cur;
    for (size_t i = 0; i < state->count; i++) {
        writer_entry *e = &state->entries[i];
        size_t nlen = strlen(e->name);

        out[cd_cur + 0] = 0x50; out[cd_cur + 1] = 0x4b; out[cd_cur + 2] = 0x01; out[cd_cur + 3] = 0x02;
        out[cd_cur + 4] = 20; out[cd_cur + 5] = 0;
        out[cd_cur + 6] = 20; out[cd_cur + 7] = 0;
        out[cd_cur + 8] = 0x00; out[cd_cur + 9] = 0x08; // UTF-8
        out[cd_cur + 10] = 0; out[cd_cur + 11] = 0;     // STORE
        out[cd_cur + 12] = (uint8_t)(dos_time); out[cd_cur + 13] = (uint8_t)(dos_time >> 8);
        out[cd_cur + 14] = (uint8_t)(dos_date); out[cd_cur + 15] = (uint8_t)(dos_date >> 8);
        out[cd_cur + 16] = (uint8_t)(e->crc); out[cd_cur + 17] = (uint8_t)(e->crc >> 8);
        out[cd_cur + 18] = (uint8_t)(e->crc >> 16); out[cd_cur + 19] = (uint8_t)(e->crc >> 24);
        out[cd_cur + 20] = (uint8_t)(e->size); out[cd_cur + 21] = (uint8_t)(e->size >> 8);
        out[cd_cur + 22] = (uint8_t)(e->size >> 16); out[cd_cur + 23] = (uint8_t)(e->size >> 24);
        out[cd_cur + 24] = (uint8_t)(e->size); out[cd_cur + 25] = (uint8_t)(e->size >> 8);
        out[cd_cur + 26] = (uint8_t)(e->size >> 16); out[cd_cur + 27] = (uint8_t)(e->size >> 24);
        out[cd_cur + 28] = (uint8_t)(nlen); out[cd_cur + 29] = (uint8_t)(nlen >> 8);
        out[cd_cur + 30] = 0; out[cd_cur + 31] = 0;
        out[cd_cur + 32] = 0; out[cd_cur + 33] = 0;
        out[cd_cur + 34] = 0; out[cd_cur + 35] = 0;
        out[cd_cur + 36] = 0; out[cd_cur + 37] = 0;
        out[cd_cur + 38] = 0; out[cd_cur + 39] = 0; out[cd_cur + 40] = 0; out[cd_cur + 41] = 0;
        out[cd_cur + 42] = (uint8_t)(e->offset); out[cd_cur + 43] = (uint8_t)(e->offset >> 8);
        out[cd_cur + 44] = (uint8_t)(e->offset >> 16); out[cd_cur + 45] = (uint8_t)(e->offset >> 24);
        memcpy(out + cd_cur + 46, e->name, nlen);
        cd_cur += 46 + nlen;
    }

    out[cd_cur + 0] = 0x50; out[cd_cur + 1] = 0x4b; out[cd_cur + 2] = 0x05; out[cd_cur + 3] = 0x06;
    out[cd_cur + 4] = 0; out[cd_cur + 5] = 0;
    out[cd_cur + 6] = 0; out[cd_cur + 7] = 0;
    out[cd_cur + 8] = (uint8_t)(state->count); out[cd_cur + 9] = (uint8_t)(state->count >> 8);
    out[cd_cur + 10] = (uint8_t)(state->count); out[cd_cur + 11] = (uint8_t)(state->count >> 8);
    out[cd_cur + 12] = (uint8_t)(cd_size); out[cd_cur + 13] = (uint8_t)(cd_size >> 8);
    out[cd_cur + 14] = (uint8_t)(cd_size >> 16); out[cd_cur + 15] = (uint8_t)(cd_size >> 24);
    out[cd_cur + 16] = (uint8_t)(cd_offset); out[cd_cur + 17] = (uint8_t)(cd_offset >> 8);
    out[cd_cur + 18] = (uint8_t)(cd_offset >> 16); out[cd_cur + 19] = (uint8_t)(cd_offset >> 24);
    out[cd_cur + 20] = 0; out[cd_cur + 21] = 0;

    *ppBuf = out;
    *pSize = total_size;
    return 1;
}

int mz_zip_writer_end(mz_zip_archive *pZip) {
    if (!pZip || !pZip->m_pState) return 0;
    zip_writer_state *state = (zip_writer_state*)pZip->m_pState;
    if (state->entries) {
        for (size_t i = 0; i < state->count; i++) {
            if (state->entries[i].name) free(state->entries[i].name);
            if (state->entries[i].data) free(state->entries[i].data);
        }
        free(state->entries);
    }
    free(state);
    pZip->m_pState = NULL;
    return 1;
}
