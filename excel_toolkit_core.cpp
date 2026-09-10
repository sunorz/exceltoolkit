/**
 * ============================================================================
 * excel_toolkit_core.cpp - OpenXML / ZIP 内存安全深度实现
 * ============================================================================
 */

#define _CRT_SECURE_NO_WARNINGS
#include "excel_toolkit_core.h"
#include "miniz.h"

#include <iostream>
#include <fstream>
#include <sstream>
#include <algorithm>
#include <chrono>
#include <iomanip>
#include <random>
#include <cstdlib>
#include <filesystem>
#include <map>
#include <set>
#include <cstring>

#ifdef _WIN32
#define WIN32_LEAN_AND_MEAN
#include <windows.h>
#endif

namespace fs = std::filesystem;

static bool ReadEntireFile(const std::string& u8path, std::vector<uint8_t>& outBuffer, std::string& err) {
    try {
#ifdef _WIN32
        int wlen = MultiByteToWideChar(CP_UTF8, 0, u8path.c_str(), -1, NULL, 0);
        if (wlen <= 0) { err = "路径编码解析失败"; return false; }
        std::wstring wpath(wlen - 1, 0);
        MultiByteToWideChar(CP_UTF8, 0, u8path.c_str(), -1, &wpath[0], wlen);
        FILE* fp = _wfopen(wpath.c_str(), L"rb");
#else
        FILE* fp = fopen(u8path.c_str(), "rb");
#endif
        if (!fp) {
            err = "无法打开文件，请确认文件存在且未被 Excel 占用锁定。";
            return false;
        }
        fseek(fp, 0, SEEK_END);
        long sz = ftell(fp);
        fseek(fp, 0, SEEK_SET);
        if (sz < 0) {
            fclose(fp);
            err = "获取文件大小失败";
            return false;
        }
        outBuffer.resize((size_t)sz);
        if (sz > 0) {
            size_t readBytes = fread(outBuffer.data(), 1, (size_t)sz, fp);
            if (readBytes != (size_t)sz) {
                fclose(fp);
                err = "读取文件内容不完整";
                return false;
            }
        }
        fclose(fp);
        return true;
    } catch (const std::exception& e) {
        err = std::string("读取文件异常: ") + e.what();
        return false;
    } catch (...) {
        err = "读取文件发生未知异常";
        return false;
    }
}

static bool WriteEntireFile(const std::string& u8path, const void* data, size_t size, std::string& err) {
    try {
#ifdef _WIN32
        int wlen = MultiByteToWideChar(CP_UTF8, 0, u8path.c_str(), -1, NULL, 0);
        if (wlen <= 0) { err = "路径编码解析失败"; return false; }
        std::wstring wpath(wlen - 1, 0);
        MultiByteToWideChar(CP_UTF8, 0, u8path.c_str(), -1, &wpath[0], wlen);
        FILE* fp = _wfopen(wpath.c_str(), L"wb");
#else
        FILE* fp = fopen(u8path.c_str(), "wb");
#endif
        if (!fp) {
            err = "无法写入目标文件，请确认目标文件未被其他程序打开且拥有写入权限。";
            return false;
        }
        if (size > 0 && data != nullptr) {
            size_t written = fwrite(data, 1, size, fp);
            if (written != size) {
                fclose(fp);
                err = "写入文件数据不完整";
                return false;
            }
        }
        fclose(fp);
        return true;
    } catch (const std::exception& e) {
        err = std::string("写入文件异常: ") + e.what();
        return false;
    } catch (...) {
        err = "写入文件发生未知异常";
        return false;
    }
}

static std::string EscapeXml(const std::string& str) {
    std::string result;
    for (char c : str) {
        if (c == '&') result += "&amp;";
        else if (c == '<') result += "&lt;";
        else if (c == '>') result += "&gt;";
        else if (c == '"') result += "&quot;";
        else if (c == '\'') result += "&apos;";
        else result += c;
    }
    return result;
}

static void InsertWorkbookProtection(std::string& wbXml, const std::string& protTag) {
    size_t wpPos = wbXml.find("<workbookProtection");
    if (wpPos != std::string::npos) {
        size_t wpEnd = wbXml.find("/>", wpPos);
        if (wpEnd != std::string::npos) {
            wbXml.erase(wpPos, wpEnd + 2 - wpPos);
        } else {
            size_t wpEndTag = wbXml.find("</workbookProtection>", wpPos);
            if (wpEndTag != std::string::npos) {
                wbXml.erase(wpPos, wpEndTag + 21 - wpPos);
            }
        }
    }
    if (protTag.empty()) return;

    const char* afterTags[] = {
        "bookViews", "sheets", "functionGroups", "externalReferences",
        "definedNames", "calcPr", "oleSize", "customWorkbookViews",
        "pivotCaches", "smartTagPr", "smartTagTypes", "webPublishing",
        "fileRecoveryPr", "webPublishObjects", "extLst"
    };

    size_t insPos = std::string::npos;
    for (const char* tag : afterTags) {
        std::string t1 = std::string("<") + tag + " ";
        std::string t2 = std::string("<") + tag + ">";
        std::string t3 = std::string("<") + tag + "/>";
        size_t p1 = wbXml.find(t1);
        size_t p2 = wbXml.find(t2);
        size_t p3 = wbXml.find(t3);
        if (p1 != std::string::npos && p1 < insPos) insPos = p1;
        if (p2 != std::string::npos && p2 < insPos) insPos = p2;
        if (p3 != std::string::npos && p3 < insPos) insPos = p3;
    }

    if (insPos != std::string::npos) {
        wbXml.insert(insPos, protTag + "\n  ");
    } else {
        size_t endWb = wbXml.rfind("</workbook>");
        if (endWb != std::string::npos) {
            wbXml.insert(endWb, protTag + "\n");
        }
    }
}

static void InsertSheetProtection(std::string& sXml, const std::string& protTag) {
    size_t spPos = sXml.find("<sheetProtection");
    if (spPos != std::string::npos) {
        size_t spEnd = sXml.find("/>", spPos);
        if (spEnd != std::string::npos) {
            sXml.erase(spPos, spEnd + 2 - spPos);
        } else {
            size_t spEndTag = sXml.find("</sheetProtection>", spPos);
            if (spEndTag != std::string::npos) {
                sXml.erase(spPos, spEndTag + 18 - spPos);
            }
        }
    }
    if (protTag.empty()) return;

    const char* afterTags[] = {
        "protectedRanges", "scenarios", "autoFilter", "sortState", "dataConsolidate",
        "customSheetViews", "mergeCells", "phoneticPr", "conditionalFormatting",
        "dataValidations", "hyperlinks", "printOptions", "pageMargins", "pageSetup",
        "headerFooter", "rowBreaks", "colBreaks", "customProperties", "cellWatches",
        "ignoredErrors", "smartTags", "drawing", "legacyDrawing", "legacyDrawingHF",
        "legacyDrawingMacro", "picture", "oleObjects", "controls", "webPublishItems",
        "tableParts", "extLst"
    };

    size_t insPos = std::string::npos;
    for (const char* tag : afterTags) {
        std::string t1 = std::string("<") + tag + " ";
        std::string t2 = std::string("<") + tag + ">";
        std::string t3 = std::string("<") + tag + "/>";
        size_t p1 = sXml.find(t1);
        size_t p2 = sXml.find(t2);
        size_t p3 = sXml.find(t3);
        if (p1 != std::string::npos && p1 < insPos) insPos = p1;
        if (p2 != std::string::npos && p2 < insPos) insPos = p2;
        if (p3 != std::string::npos && p3 < insPos) insPos = p3;
    }

    if (insPos != std::string::npos) {
        sXml.insert(insPos, protTag + "\n  ");
    } else {
        size_t sdEnd = sXml.find("</sheetData>");
        if (sdEnd != std::string::npos) {
            sXml.insert(sdEnd + 12, "\n  " + protTag);
        } else {
            size_t sdSelfClose = sXml.find("<sheetData/>");
            if (sdSelfClose != std::string::npos) {
                sXml.insert(sdSelfClose + 12, "\n  " + protTag);
            } else {
                size_t endWs = sXml.rfind("</worksheet>");
                if (endWs != std::string::npos) {
                    sXml.insert(endWs, protTag + "\n");
                }
            }
        }
    }
}

static void InsertWorksheetDrawing(std::string& sXml, const std::string& drawingTag) {
    if (sXml.find("<drawing ") != std::string::npos || sXml.find("<drawing>") != std::string::npos) {
        return;
    }

    const char* drawingAfterTags[] = {
        "legacyDrawing", "legacyDrawingHF", "legacyDrawingMacro", "picture",
        "oleObjects", "controls", "webPublishItems", "tableParts", "extLst"
    };
    size_t insPos = std::string::npos;
    for (const char* tag : drawingAfterTags) {
        std::string t1 = std::string("<") + tag + " ";
        std::string t2 = std::string("<") + tag + ">";
        std::string t3 = std::string("<") + tag + "/>";
        size_t p1 = sXml.find(t1);
        size_t p2 = sXml.find(t2);
        size_t p3 = sXml.find(t3);
        if (p1 != std::string::npos && p1 < insPos) insPos = p1;
        if (p2 != std::string::npos && p2 < insPos) insPos = p2;
        if (p3 != std::string::npos && p3 < insPos) insPos = p3;
    }

    if (insPos != std::string::npos) {
        sXml.insert(insPos, drawingTag + "\n  ");
    } else {
        size_t endWs = sXml.rfind("</worksheet>");
        if (endWs != std::string::npos) {
            sXml.insert(endWs, "  " + drawingTag + "\n");
        }
    }
}

static void UnlockStylesXml(std::string& stXml) {
    size_t xfsPos = stXml.find("<cellXfs");
    if (xfsPos == std::string::npos) return;
    size_t xfsEnd = stXml.find("</cellXfs>", xfsPos);
    if (xfsEnd == std::string::npos) return;

    std::string cellXfsContent = stXml.substr(xfsPos, xfsEnd + 10 - xfsPos);
    std::string newCellXfs;
    newCellXfs.reserve(cellXfsContent.length() + 1024);

    size_t cur = 0;
    while (cur < cellXfsContent.length()) {
        size_t xfStart = cellXfsContent.find("<xf ", cur);
        if (xfStart == std::string::npos) {
            newCellXfs.append(cellXfsContent.substr(cur));
            break;
        }

        newCellXfs.append(cellXfsContent.substr(cur, xfStart - cur));

        size_t openTagEnd = cellXfsContent.find(">", xfStart);
        if (openTagEnd == std::string::npos) {
            newCellXfs.append(cellXfsContent.substr(xfStart));
            break;
        }

        bool isSelfClosing = (cellXfsContent[openTagEnd - 1] == '/');
        std::string openTag = cellXfsContent.substr(xfStart, openTagEnd + 1 - xfStart);

        size_t apPos = openTag.find("applyProtection=\"");
        if (apPos != std::string::npos) {
            size_t valStart = apPos + 17;
            size_t quoteEnd = openTag.find("\"", valStart);
            if (quoteEnd != std::string::npos) {
                openTag.replace(valStart, quoteEnd - valStart, "1");
            }
        } else {
            openTag.insert(4, "applyProtection=\"1\" ");
        }

        if (isSelfClosing) {
            size_t slashPos = openTag.rfind("/>");
            if (slashPos != std::string::npos) {
                openTag.replace(slashPos, 2, "><protection locked=\"0\"/></xf>");
            }
            newCellXfs.append(openTag);
            cur = openTagEnd + 1;
        } else {
            size_t xfCloseTag = cellXfsContent.find("</xf>", openTagEnd);
            if (xfCloseTag == std::string::npos) {
                newCellXfs.append(openTag);
                cur = openTagEnd + 1;
                continue;
            }

            std::string body = cellXfsContent.substr(openTagEnd + 1, xfCloseTag - (openTagEnd + 1));
            size_t pPos = body.find("<protection");
            if (pPos != std::string::npos) {
                size_t pClose = body.find(">", pPos);
                if (pClose != std::string::npos) {
                    std::string protTag = body.substr(pPos, pClose + 1 - pPos);
                    size_t lkPos = protTag.find("locked=\"");
                    if (lkPos != std::string::npos) {
                        size_t lkVal = lkPos + 8;
                        size_t lkEnd = protTag.find("\"", lkVal);
                        if (lkEnd != std::string::npos) {
                            protTag.replace(lkVal, lkEnd - lkVal, "0");
                        }
                    } else {
                        protTag.insert(11, " locked=\"0\"");
                    }
                    body.replace(pPos, pClose + 1 - pPos, protTag);
                }
            } else {
                size_t alignPos = body.find("<alignment");
                if (alignPos != std::string::npos) {
                    size_t alignEnd = body.find("</alignment>", alignPos);
                    if (alignEnd != std::string::npos) {
                        body.insert(alignEnd + 12, "<protection locked=\"0\"/>");
                    } else {
                        size_t alignSelfClose = body.find("/>", alignPos);
                        if (alignSelfClose != std::string::npos) {
                            body.insert(alignSelfClose + 2, "<protection locked=\"0\"/>");
                        } else {
                            body.insert(0, "<protection locked=\"0\"/>");
                        }
                    }
                } else {
                    size_t extPos = body.find("<extLst");
                    if (extPos != std::string::npos) {
                        body.insert(extPos, "<protection locked=\"0\"/>");
                    } else {
                        body.insert(0, "<protection locked=\"0\"/>");
                    }
                }
            }

            newCellXfs.append(openTag);
            newCellXfs.append(body);
            newCellXfs.append("</xf>");
            cur = xfCloseTag + 5;
        }
    }

    stXml.replace(xfsPos, xfsEnd + 10 - xfsPos, newCellXfs);
}

std::string ExcelToolkitCore::GenerateRandomHex8() {
    std::random_device rd;
    std::mt19937 gen(rd());
    std::uniform_int_distribution<uint32_t> dis(0, 0xFFFFFFFF);
    uint32_t val = dis(gen);
    std::stringstream ss;
    ss << std::hex << std::setw(8) << std::setfill('0') << val;
    return ss.str();
}

std::string ExcelToolkitCore::GetIsoTimestampNow() {
    auto now = std::chrono::system_clock::now();
    auto in_time_t = std::chrono::system_clock::to_time_t(now);
    std::time_t cst_time = in_time_t + 8 * 3600;
    std::tm* gmt = std::gmtime(&cst_time);
    char buf[64] = { 0 };
    if (gmt) {
        std::strftime(buf, sizeof(buf), "%Y-%m-%d %H:%M:%S (CST)", gmt);
        return std::string(buf);
    }
    return "2026-09-09 15:00:00 (CST)";
}

uint16_t ExcelToolkitCore::ComputeExcelPasswordHash(const std::string& password) {
    if (password.empty()) return 0;
    uint16_t hash = 0;
    int len = (int)password.length();
    for (int i = len - 1; i >= 0; --i) {
        uint8_t c = (uint8_t)password[i];
        hash = ((hash >> 14) & 0x01) | ((hash << 1) & 0x7FFE);
        hash ^= c;
    }
    hash = ((hash >> 14) & 0x01) | ((hash << 1) & 0x7FFE);
    hash ^= (uint16_t)len;
    hash ^= 0xCE4B;
    return hash;
}

bool ExcelToolkitCore::CheckSofficeAvailable() {
#ifdef _WIN32
    int ret = system("where soffice >nul 2>nul");
    return ret == 0;
#else
    int ret = system("which soffice >/dev/null 2>&1");
    return ret == 0;
#endif
}

std::string ExcelToolkitCore::ConvertWithSoffice(const std::string& srcPath, const std::string& targetExt, const std::string& workDir) {
    std::string cmd = "soffice --headless --norestore --convert-to " + targetExt + " --outdir \"" + workDir + "\" \"" + srcPath + "\"";
    int ret = system(cmd.c_str());
    if (ret != 0) {
        throw std::runtime_error("soffice 转换命令执行失败");
    }
    fs::path p = fs::u8path(srcPath);
    std::string stem = p.stem().u8string();
    fs::path outP = fs::u8path(workDir) / fs::u8path(stem + "." + targetExt);
    if (!fs::exists(outP)) {
        throw std::runtime_error("soffice 转换后未找到输出文件: " + outP.u8string());
    }
    return outP.u8string();
}

struct MemoryZipPackage {
    std::map<std::string, std::string> files;

    bool LoadFromBuffer(const std::vector<uint8_t>& buffer, std::string& err) {
        if (buffer.size() < 22) {
            err = "文件过小，非有效 Excel 压缩容器。";
            return false;
        }

        mz_zip_archive zip;
        memset(&zip, 0, sizeof(zip));
        if (!mz_zip_reader_init_mem(&zip, buffer.data(), buffer.size(), 0)) {
            err = "无法以 ZIP 容器解析该 XLSX 文件（可能损坏或非标准 xlsx 格式）。";
            return false;
        }

        uint32_t numFiles = mz_zip_reader_get_num_files(&zip);
        for (uint32_t i = 0; i < numFiles; ++i) {
            mz_zip_archive_file_stat stat;
            if (!mz_zip_reader_file_stat(&zip, i, &stat)) continue;
            if (stat.m_is_directory) continue;

            size_t size = 0;
            void* pData = mz_zip_reader_extract_to_heap(&zip, i, &size, 0);
            if (pData) {
                files[stat.m_filename] = std::string((char*)pData, size);
                mz_free(pData);
            }
        }
        mz_zip_reader_end(&zip);
        return true;
    }

    bool SaveToFile(const std::string& outputPath, std::string& err) {
        mz_zip_archive zip;
        memset(&zip, 0, sizeof(zip));
        if (!mz_zip_writer_init_heap(&zip, 0, 0)) {
            err = "初始化内存 ZIP 压缩器失败";
            return false;
        }

        for (const auto& kv : files) {
            mz_zip_writer_add_mem(&zip, kv.first.c_str(), kv.second.data(), kv.second.size(), MZ_DEFAULT_COMPRESSION);
        }

        void* pBuf = nullptr;
        size_t bufSize = 0;
        if (!mz_zip_writer_finalize_heap_archive(&zip, &pBuf, &bufSize)) {
            mz_zip_writer_end(&zip);
            err = "打包 XLSX 数据失败";
            return false;
        }

        bool writeOk = WriteEntireFile(outputPath, pBuf, bufSize, err);
        mz_zip_writer_end(&zip);
        return writeOk;
    }
};

bool ExcelToolkitCore::ProcessExcelFile(
    const std::string& inputPath,
    const std::string& outputPath,
    const ProcessOptions& options,
    std::string& outError
) {
    try {
        std::vector<uint8_t> inBuffer;
        if (!ReadEntireFile(inputPath, inBuffer, outError)) {
            return false;
        }

        MemoryZipPackage pkg;
        if (!pkg.LoadFromBuffer(inBuffer, outError)) {
            return false;
        }

        if (pkg.files.find("xl/workbook.xml") == pkg.files.end()) {
            outError = "未找到核心文件 xl/workbook.xml，该文件可能不是标准的 XLSX 文件。";
            return false;
        }

        std::string& wbXml = pkg.files["xl/workbook.xml"];
        std::string& relsXml = pkg.files["xl/_rels/workbook.xml.rels"];
        std::string& ctXml = pkg.files["[Content_Types].xml"];

        if (options.unprotectAll) {
            size_t wpPos = wbXml.find("<workbookProtection");
            if (wpPos != std::string::npos) {
                size_t wpEnd = wbXml.find("/>", wpPos);
                if (wpEnd != std::string::npos) {
                    wbXml.erase(wpPos, wpEnd + 2 - wpPos);
                }
            }

            for (auto& kv : pkg.files) {
                if (kv.first.rfind("xl/worksheets/sheet", 0) == 0 && kv.first.find(".xml") != std::string::npos) {
                    std::string& sXml = kv.second;
                    size_t spPos = sXml.find("<sheetProtection");
                    if (spPos != std::string::npos) {
                        size_t spEnd = sXml.find("/>", spPos);
                        if (spEnd != std::string::npos) {
                            sXml.erase(spPos, spEnd + 2 - spPos);
                        }
                    }
                }
            }
        } else {
            // 1. 锁定工作簿结构 (严格遵循 OpenXML XSD Sequence：位于 bookViews, sheets 之前)
            if (options.lockWorkbookStructure) {
                std::string wbProt;
                if (!options.password.empty()) {
                    uint16_t wbPwHash = ComputeExcelPasswordHash(options.password);
                    std::stringstream wbHex;
                    wbHex << std::hex << std::uppercase << std::setw(4) << std::setfill('0') << wbPwHash;
                    wbProt = "<workbookProtection workbookPassword=\"" + wbHex.str() + "\" lockStructure=\"1\"/>";
                } else {
                    wbProt = "<workbookProtection lockStructure=\"1\"/>";
                }
                InsertWorkbookProtection(wbXml, wbProt);
            }

            // 2. 解锁所有单元格样式 (为实体可见水印模式放开修改权限)
            if (options.unlockAllCells) {
                if (pkg.files.find("xl/styles.xml") != pkg.files.end()) {
                    UnlockStylesXml(pkg.files["xl/styles.xml"]);
                }
            }

            // 3. 工作表保护 (针对水印保护 / 保护整个文件不受更改)
            // 当勾选添加实体可见水印时：必须对水印对象进行锁定保护（objects="1"），
            // 同时放行所有单元格的选择与编辑（selectUnlockedCells="1"，配合单元格已设置 locked="0"），
            // 确保全屏实体水印不可被选中、不可被修改、不可被删除，而单元格内容完全可随意修改！
            if (options.protectSheetContent || options.addVisibleWatermark) {
                std::string protTag;
                std::string pwAttr = "";
                if (!options.password.empty()) {
                    uint16_t pwHash = ComputeExcelPasswordHash(options.password);
                    std::stringstream hexHash;
                    hexHash << std::hex << std::uppercase << std::setw(4) << std::setfill('0') << pwHash;
                    pwAttr = " password=\"" + hexHash.str() + "\"";
                }

                if (options.protectSheetContent && !options.unlockAllCells) {
                    // 全表锁定模式：严格禁止修改内容与对象
                    protTag = "<sheetProtection" + pwAttr + " sheet=\"1\" objects=\"1\" scenarios=\"1\" selectLockedCells=\"1\" selectUnlockedCells=\"1\"/>";
                } else {
                    // 水印保护模式：锁定图形水印对象，放行单元格编辑与输入
                    protTag = "<sheetProtection" + pwAttr + " sheet=\"1\" objects=\"1\" scenarios=\"0\" selectLockedCells=\"1\" selectUnlockedCells=\"1\"/>";
                }

                for (auto& kv : pkg.files) {
                    if (kv.first.rfind("xl/worksheets/sheet", 0) == 0 && kv.first.find(".xml") != std::string::npos) {
                        if (kv.first.find("sheet_sysdata_cfg") != std::string::npos) continue;
                        InsertSheetProtection(kv.second, protTag);
                    }
                }
            }

            // 4. 添加实体可见水印 (全表平铺矩阵显示，文字半透明倾斜，背景无填充不挡鼠标点击)
            if (options.addVisibleWatermark) {
                std::string entityWm = options.visibleWatermarkText.empty() ? "内部资料 严禁外传" : options.visibleWatermarkText;
                
                int maxDrawingId = 0;
                for (const auto& kv : pkg.files) {
                    if (kv.first.rfind("xl/drawings/drawing", 0) == 0) {
                        size_t numStart = 19;
                        size_t dot = kv.first.find(".xml", numStart);
                        if (dot != std::string::npos) {
                            try {
                                int idVal = std::stoi(kv.first.substr(numStart, dot - numStart));
                                if (idVal > maxDrawingId) maxDrawingId = idVal;
                            } catch (...) {}
                        }
                    }
                }

                for (auto& kv : pkg.files) {
                    if (kv.first.rfind("xl/worksheets/sheet", 0) == 0 && kv.first.find(".xml") != std::string::npos) {
                        if (kv.first.find("sheet_sysdata_cfg") != std::string::npos) continue;
                        std::string sheetPath = kv.first;
                        std::string& sXml = kv.second;

                        size_t lastSlash = sheetPath.find_last_of('/');
                        std::string sheetFileName = sheetPath.substr(lastSlash + 1);

                        size_t dPos = sXml.find("<drawing r:id=\"");
                        std::string drawRelId = "";
                        if (dPos != std::string::npos) {
                            size_t rEnd = sXml.find("\"", dPos + 15);
                            if (rEnd != std::string::npos) {
                                drawRelId = sXml.substr(dPos + 15, rEnd - (dPos + 15));
                            }
                        }

                        std::string sheetRelsPath = "xl/worksheets/_rels/" + sheetFileName + ".rels";
                        int drawingId = 0;
                        
                        if (!drawRelId.empty() && pkg.files.find(sheetRelsPath) != pkg.files.end()) {
                            std::string& sRels = pkg.files[sheetRelsPath];
                            size_t rPos = sRels.find("Id=\"" + drawRelId + "\"");
                            if (rPos != std::string::npos) {
                                size_t tPos = sRels.find("Target=\"", rPos);
                                if (tPos != std::string::npos) {
                                    size_t tEnd = sRels.find("\"", tPos + 8);
                                    if (tEnd != std::string::npos) {
                                        std::string target = sRels.substr(tPos + 8, tEnd - (tPos + 8));
                                        size_t numStart = target.find("drawing");
                                        if (numStart != std::string::npos) {
                                            numStart += 7;
                                            size_t dot = target.find(".xml", numStart);
                                            if (dot != std::string::npos) {
                                                try {
                                                    drawingId = std::stoi(target.substr(numStart, dot - numStart));
                                                } catch (...) {}
                                            }
                                        }
                                    }
                                }
                            }
                        }

                        if (drawingId == 0) {
                            maxDrawingId++;
                            drawingId = maxDrawingId;
                            drawRelId = "rIdWatermark" + std::to_string(drawingId);
                            
                            std::string relEntry = "<Relationship Id=\"" + drawRelId + "\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing\" Target=\"../drawings/drawing" + std::to_string(drawingId) + ".xml\"/>";
                            if (pkg.files.find(sheetRelsPath) != pkg.files.end()) {
                                std::string& sRels = pkg.files[sheetRelsPath];
                                size_t rEnd = sRels.find("</Relationships>");
                                if (rEnd != std::string::npos) {
                                    sRels.insert(rEnd, "  " + relEntry + "\n");
                                }
                            } else {
                                pkg.files[sheetRelsPath] = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\n  " + relEntry + "\n</Relationships>";
                            }

                            std::string drawingTag = "<drawing r:id=\"" + drawRelId + "\"/>";
                            InsertWorksheetDrawing(sXml, drawingTag);
                        }

                        std::string drawingXmlPath = "xl/drawings/drawing" + std::to_string(drawingId) + ".xml";
                        std::string& dXml = pkg.files[drawingXmlPath];

                        if (dXml.empty()) {
                            dXml = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<xdr:wsDr xmlns:xdr=\"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\">\n</xdr:wsDr>";
                        }
                        
                        size_t wsDrEnd = dXml.find("</xdr:wsDr>");
                        if (wsDrEnd != std::string::npos) {
                            // 全屏平铺矩阵：在 0~80行、0~30列范围内生成均匀分布的倾斜半透明水印文本框
                            // 背景为 <a:noFill/>，边框为无边框，绝不拦截鼠标对单元格的点击和双击编辑！
                            std::stringstream ssShape;
                            int shapeCounter = 0;
                            for (int r = 0; r <= 60; r += 15) {
                                for (int c = 0; c <= 24; c += 8) {
                                    int shapeId = 990000 + drawingId * 100 + shapeCounter++;
                                    ssShape << "  <xdr:twoCellAnchor editAs=\"oneCell\">\n"
                                            << "    <xdr:from><xdr:col>" << c << "</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>" << r << "</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>\n"
                                            << "    <xdr:to><xdr:col>" << (c + 7) << "</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>" << (r + 10) << "</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>\n"
                                            << "    <xdr:sp macro=\"\" textlink=\"\">\n"
                                            << "      <xdr:nvSpPr>\n"
                                            << "        <xdr:cNvPr id=\"" << shapeId << "\" name=\"Watermark_" << shapeId << "\"/>\n"
                                            << "        <xdr:cNvSpPr txBox=\"1\">\n"
                                            << "          <a:spLocks noSelect=\"1\" noMove=\"1\" noResize=\"1\" noEditPoints=\"1\" noAdjustHandles=\"1\" noChangeArrowheads=\"1\" noChangeShapeType=\"1\" noGrouping=\"1\"/>\n"
                                            << "        </xdr:cNvSpPr>\n"
                                            << "      </xdr:nvSpPr>\n"
                                            << "      <xdr:spPr>\n"
                                            << "        <a:xfrm rot=\"-1800000\"><a:off x=\"0\" y=\"0\"/><a:ext cx=\"0\" cy=\"0\"/></a:xfrm>\n"
                                            << "        <a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom>\n"
                                            << "        <a:noFill/>\n"
                                            << "        <a:ln><a:noFill/></a:ln>\n"
                                            << "      </xdr:spPr>\n"
                                            << "      <xdr:txBody>\n"
                                            << "        <a:bodyPr vertOverflow=\"clip\" horzOverflow=\"clip\" wrap=\"none\" rtlCol=\"0\" anchor=\"ctr\"/>\n"
                                            << "        <a:lstStyle/>\n"
                                            << "        <a:p>\n"
                                            << "          <a:pPr algn=\"ctr\"/>\n"
                                            << "          <a:r>\n"
                                            << "            <a:rPr lang=\"zh-CN\" sz=\"2400\" b=\"1\">\n"
                                            << "              <a:solidFill><a:srgbClr val=\"909090\"><a:alpha val=\"22000\"/></a:srgbClr></a:solidFill>\n"
                                            << "              <a:latin typeface=\"Microsoft YaHei\"/>\n"
                                            << "              <a:ea typeface=\"Microsoft YaHei\"/>\n"
                                            << "            </a:rPr>\n"
                                            << "            <a:t>" << EscapeXml(entityWm) << "</a:t>\n"
                                            << "          </a:r>\n"
                                            << "        </a:p>\n"
                                            << "      </xdr:txBody>\n"
                                            << "    </xdr:sp>\n"
                                            << "    <xdr:clientData fLocksWithSheet=\"1\" fPrintsWithSheet=\"1\"/>\n"
                                            << "  </xdr:twoCellAnchor>\n";
                                }
                            }
                            dXml.insert(wsDrEnd, ssShape.str());
                        }

                        if (ctXml.find("/" + drawingXmlPath) == std::string::npos) {
                            std::string ctTag = "<Override PartName=\"/" + drawingXmlPath + "\" ContentType=\"application/vnd.openxmlformats-officedocument.drawing+xml\"/>";
                            size_t ctEnd = ctXml.find("</Types>");
                            if (ctEnd != std::string::npos) {
                                ctXml.insert(ctEnd, "  " + ctTag + "\n");
                            }
                        }
                    }
                }
            }

            if (options.addWatermark) {
                std::string effectiveSheetName = options.sheetName.empty() ? HIDDEN_SHEET_NAME_CONST : options.sheetName;
                std::string effectiveWmText = options.watermarkText.empty() ? "绝密外发审计凭证" : options.watermarkText;
                
                std::string effectiveIssuer = options.issuer.empty() ? "ADMIN" : options.issuer;
                std::string effectiveTraceId = options.traceId;
                if (effectiveTraceId.empty()) {
                    auto nowSec = std::chrono::duration_cast<std::chrono::seconds>(std::chrono::system_clock::now().time_since_epoch()).count();
                    effectiveTraceId = effectiveIssuer + "-" + std::to_string(nowSec);
                }
                std::string nowIso = GetIsoTimestampNow();

                std::string targetSheetFile = "worksheets/sheet_sysdata_cfg.xml";
                std::string targetSheetRelId = "rIdSysDataCfg";
                int targetSheetId = 9988;

                size_t oldSheetPos = wbXml.find("name=\"" + effectiveSheetName + "\"");
                if (oldSheetPos != std::string::npos) {
                    size_t tagStart = wbXml.rfind("<sheet", oldSheetPos);
                    size_t tagEnd = wbXml.find("/>", oldSheetPos);
                    if (tagEnd == std::string::npos) tagEnd = wbXml.find("</sheet>", oldSheetPos) + 8;
                    else tagEnd += 2;
                    if (tagStart != std::string::npos && tagEnd != std::string::npos) {
                        wbXml.erase(tagStart, tagEnd - tagStart);
                    }
                }

                std::stringstream ssSheet;
                ssSheet << "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
                        << "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">\n"
                        << "  <dimension ref=\"A1:B3\"/>\n"
                        << "  <sheetViews><sheetView tabSelected=\"0\" workbookViewId=\"0\"/></sheetViews>\n"
                        << "  <sheetData>\n"
                        << "    <row r=\"1\"><c r=\"A1\" t=\"inlineStr\"><is><t>溯源编号</t></is></c><c r=\"B1\" t=\"inlineStr\"><is><t>" << effectiveTraceId << "</t></is></c></row>\n"
                        << "    <row r=\"2\"><c r=\"A2\" t=\"inlineStr\"><is><t>水印文字</t></is></c><c r=\"B2\" t=\"inlineStr\"><is><t>" << effectiveWmText << "</t></is></c></row>\n"
                        << "    <row r=\"3\"><c r=\"A3\" t=\"inlineStr\"><is><t>生成时间</t></is></c><c r=\"B3\" t=\"inlineStr\"><is><t>" << nowIso << "</t></is></c></row>\n"
                        << "  </sheetData>\n"
                        << "</worksheet>";

                pkg.files["xl/" + targetSheetFile] = ssSheet.str();

                std::string sheetTag = "<sheet name=\"" + effectiveSheetName + "\" sheetId=\"" + std::to_string(targetSheetId) + "\" state=\"veryHidden\" r:id=\"" + targetSheetRelId + "\"/>";
                size_t sheetsEnd = wbXml.find("</sheets>");
                if (sheetsEnd != std::string::npos) {
                    wbXml.insert(sheetsEnd, "  " + sheetTag + "\n");
                } else {
                    wbXml = "<workbook><sheets>" + sheetTag + "</sheets></workbook>";
                }

                if (relsXml.find("Id=\"" + targetSheetRelId + "\"") == std::string::npos) {
                    std::string relTag = "<Relationship Id=\"" + targetSheetRelId + "\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"" + targetSheetFile + "\"/>";
                    size_t relsEnd = relsXml.find("</Relationships>");
                    if (relsEnd != std::string::npos) {
                        relsXml.insert(relsEnd, "  " + relTag + "\n");
                    }
                }

                if (ctXml.find("/xl/" + targetSheetFile) == std::string::npos) {
                    std::string ctTag = "<Override PartName=\"/xl/" + targetSheetFile + "\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>";
                    size_t ctEnd = ctXml.find("</Types>");
                    if (ctEnd != std::string::npos) {
                        ctXml.insert(ctEnd, "  " + ctTag + "\n");
                    }
                }

                if (pkg.files.find("docProps/app.xml") != pkg.files.end()) {
                    std::string& appXml = pkg.files["docProps/app.xml"];

                    std::vector<std::string> sheetNames;
                    size_t searchPos = 0;
                    while (true) {
                        size_t sheetTagPos = wbXml.find("<sheet ", searchPos);
                        if (sheetTagPos == std::string::npos) break;
                        size_t tagClose = wbXml.find("/>", sheetTagPos);
                        if (tagClose == std::string::npos) break;
                        std::string tagContent = wbXml.substr(sheetTagPos, tagClose - sheetTagPos);
                        size_t namePos = tagContent.find("name=\"");
                        if (namePos != std::string::npos) {
                            size_t nameStart = namePos + 6;
                            size_t nameEnd = tagContent.find("\"", nameStart);
                            if (nameEnd != std::string::npos) {
                                sheetNames.push_back(tagContent.substr(nameStart, nameEnd - nameStart));
                            }
                        }
                        searchPos = tagClose + 2;
                    }

                    if (!sheetNames.empty()) {
                        std::stringstream ssTitles;
                        for (const auto& nm : sheetNames) {
                            ssTitles << "<vt:lpstr>" << EscapeXml(nm) << "</vt:lpstr>";
                        }
                        std::string newTitlesVector = "<vt:vector size=\"" + std::to_string(sheetNames.size())
                            + "\" baseType=\"lpstr\">" + ssTitles.str() + "</vt:vector>";

                        size_t topPos = appXml.find("<TitlesOfParts>");
                        if (topPos != std::string::npos) {
                            size_t topEnd = appXml.find("</TitlesOfParts>");
                            if (topEnd != std::string::npos) {
                                size_t vecStart = appXml.find("<vt:vector", topPos);
                                if (vecStart != std::string::npos && vecStart < topEnd) {
                                    appXml.replace(vecStart, topEnd - vecStart, newTitlesVector);
                                }
                            }
                        } else {
                            size_t propEnd = appXml.find("</Properties>");
                            if (propEnd != std::string::npos) {
                                appXml.insert(propEnd, "<TitlesOfParts>" + newTitlesVector + "</TitlesOfParts>");
                            }
                        }

                        size_t hpPos = appXml.find("<HeadingPairs>");
                        if (hpPos != std::string::npos) {
                            size_t hpEnd = appXml.find("</HeadingPairs>");
                            if (hpEnd != std::string::npos) {
                                size_t i4Pos = appXml.find("<vt:i4>", hpPos);
                                if (i4Pos != std::string::npos && i4Pos < hpEnd) {
                                    size_t i4End = appXml.find("</vt:i4>", i4Pos);
                                    if (i4End != std::string::npos) {
                                        std::string newI4 = "<vt:i4>" + std::to_string(sheetNames.size()) + "</vt:i4>";
                                        appXml.replace(i4Pos, i4End + 8 - i4Pos, newI4);
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }

        std::string finalOutput = outputPath.empty() ? inputPath : outputPath;
        return pkg.SaveToFile(finalOutput, outError);
    } catch (const std::exception& e) {
        outError = std::string("处理异常: ") + e.what();
        return false;
    } catch (...) {
        outError = "处理发生未知内部异常";
        return false;
    }
}

bool ExcelToolkitCore::InspectWatermarkFile(
    const std::string& filePath,
    InspectWatermarkInfo& outInfo,
    std::string& outError
) {
    try {
        std::vector<uint8_t> inBuffer;
        if (!ReadEntireFile(filePath, inBuffer, outError)) {
            return false;
        }

        MemoryZipPackage pkg;
        if (!pkg.LoadFromBuffer(inBuffer, outError)) {
            return false;
        }

        if (pkg.files.find("xl/workbook.xml") == pkg.files.end()) {
            outError = "缺少核心结构 xl/workbook.xml";
            return false;
        }

        const std::string& wbXml = pkg.files["xl/workbook.xml"];
        std::string relsXml = "";
        if (pkg.files.find("xl/_rels/workbook.xml.rels") != pkg.files.end()) {
            relsXml = pkg.files["xl/_rels/workbook.xml.rels"];
        }

        if (wbXml.find("lockStructure=\"1\"") != std::string::npos || wbXml.find("lockStructure=\"true\"") != std::string::npos) {
            outInfo.isWorkbookLocked = true;
        }

        std::map<std::string, std::string> rIdToTarget;
        size_t rPos = 0;
        while ((rPos = relsXml.find("<Relationship ", rPos)) != std::string::npos) {
            size_t rEnd = relsXml.find("/>", rPos);
            if (rEnd == std::string::npos) break;
            std::string rTag = relsXml.substr(rPos, rEnd - rPos);

            size_t idPos = rTag.find("Id=\"");
            size_t targetPos = rTag.find("Target=\"");
            if (idPos != std::string::npos && targetPos != std::string::npos) {
                std::string id = rTag.substr(idPos + 4, rTag.find("\"", idPos + 4) - (idPos + 4));
                std::string target = rTag.substr(targetPos + 8, rTag.find("\"", targetPos + 8) - (targetPos + 8));
                rIdToTarget[id] = target;
            }
            rPos = rEnd;
        }

        std::vector<std::string> sharedStrings;
        if (pkg.files.find("xl/sharedStrings.xml") != pkg.files.end()) {
            const std::string& ssXml = pkg.files["xl/sharedStrings.xml"];
            size_t siPos = 0;
            while ((siPos = ssXml.find("<si", siPos)) != std::string::npos) {
                size_t siEnd = ssXml.find("</si>", siPos);
                if (siEnd == std::string::npos) break;
                std::string siChunk = ssXml.substr(siPos, siEnd - siPos);
                std::string tCombined = "";
                size_t tp = 0;
                while ((tp = siChunk.find("<t", tp)) != std::string::npos) {
                    size_t valStart = siChunk.find(">", tp);
                    if (valStart == std::string::npos) break;
                    valStart++;
                    size_t valEnd = siChunk.find("</t>", valStart);
                    if (valEnd != std::string::npos) {
                        tCombined += siChunk.substr(valStart, valEnd - valStart);
                        tp = valEnd + 4;
                    } else {
                        break;
                    }
                }
                sharedStrings.push_back(tCombined);
                siPos = siEnd + 5;
            }
        }

        auto findSecPattern = [](const std::string& str) -> std::string {
            size_t pos = str.find("SEC-");
            if (pos == std::string::npos) return "";
            size_t endPos = pos + 4;
            while (endPos < str.size() && (isalnum((unsigned char)str[endPos]) || str[endPos] == '-' || str[endPos] == '_')) {
                endPos++;
            }
            return str.substr(pos, endPos - pos);
        };

        size_t sPos = 0;
        std::string hiddenSheetXmlPath = "";

        while ((sPos = wbXml.find("<sheet ", sPos)) != std::string::npos) {
            size_t sEnd = wbXml.find(">", sPos);
            if (sEnd == std::string::npos) break;

            std::string tag = wbXml.substr(sPos, sEnd - sPos);

            std::string name = "";
            size_t nPos = tag.find("name=\"");
            if (nPos != std::string::npos) {
                name = tag.substr(nPos + 6, tag.find("\"", nPos + 6) - (nPos + 6));
            }

            std::string state = "visible";
            size_t stPos = tag.find("state=\"");
            if (stPos != std::string::npos) {
                state = tag.substr(stPos + 7, tag.find("\"", stPos + 7) - (stPos + 7));
            }

            std::string rId = "";
            size_t ridPos = tag.find("r:id=\"");
            if (ridPos != std::string::npos) {
                rId = tag.substr(ridPos + 6, tag.find("\"", ridPos + 6) - (ridPos + 6));
            }

            SheetMeta sm;
            sm.name = name;
            sm.state = state;
            outInfo.sheets.push_back(sm);

            if (name == HIDDEN_SHEET_NAME_CONST || state == "veryHidden") {
                outInfo.hasHiddenSheet = true;
                if (rIdToTarget.find(rId) != rIdToTarget.end()) {
                    std::string t = rIdToTarget[rId];
                    while (!t.empty() && (t[0] == '/' || t[0] == '\\')) t.erase(0, 1);
                    if (t.rfind("./", 0) == 0) t.erase(0, 2);

                    if (t.rfind("xl/", 0) == 0) {
                        hiddenSheetXmlPath = t;
                    } else if (t.rfind("worksheets/", 0) == 0) {
                        hiddenSheetXmlPath = "xl/" + t;
                    } else {
                        hiddenSheetXmlPath = "xl/worksheets/" + t;
                    }
                } else {
                    hiddenSheetXmlPath = "xl/worksheets/sheet_sysdata_cfg.xml";
                }
            }

            sPos = sEnd;
        }

        auto extractCellText = [&](const std::string& ref, const std::string& xml) -> std::string {
            std::string search1 = "<c r=\"" + ref + "\"";
            std::string search2 = " r=\"" + ref + "\"";
            size_t cPos = xml.find(search1);
            if (cPos == std::string::npos) {
                size_t cur = 0;
                while ((cur = xml.find("<c ", cur)) != std::string::npos) {
                    size_t closeTag = xml.find(">", cur);
                    if (closeTag != std::string::npos) {
                        std::string cTag = xml.substr(cur, closeTag - cur);
                        if (cTag.find(search2) != std::string::npos) {
                            cPos = cur;
                            break;
                        }
                    }
                    cur += 3;
                }
            }
            if (cPos == std::string::npos) return "";

            size_t cEnd = xml.find("</c>", cPos);
            if (cEnd == std::string::npos) cEnd = xml.find("/>", cPos);
            if (cEnd == std::string::npos) return "";

            std::string cContent = xml.substr(cPos, cEnd - cPos + 4);
            if (cContent.find("t=\"s\"") != std::string::npos) {
                size_t vPos = cContent.find("<v>");
                if (vPos != std::string::npos) {
                    size_t vEnd = cContent.find("</v>", vPos);
                    if (vEnd != std::string::npos) {
                        std::string idxStr = cContent.substr(vPos + 3, vEnd - (vPos + 3));
                        try {
                            size_t idx = std::stoul(idxStr);
                            if (idx < sharedStrings.size()) {
                                return sharedStrings[idx];
                            }
                        } catch (...) {}
                    }
                }
            }
            
            size_t tPos = cContent.find("<t>");
            if (tPos == std::string::npos) tPos = cContent.find("<t ");
            if (tPos != std::string::npos) {
                size_t tValStart = cContent.find(">", tPos) + 1;
                size_t tEnd = cContent.find("</t>", tValStart);
                if (tEnd != std::string::npos) {
                    return cContent.substr(tValStart, tEnd - tValStart);
                }
            }

            size_t vPos = cContent.find("<v>");
            if (vPos != std::string::npos) {
                size_t vEnd = cContent.find("</v>", vPos);
                if (vEnd != std::string::npos) {
                    return cContent.substr(vPos + 3, vEnd - (vPos + 3));
                }
            }
            
            return "";
        };

        auto parseCommentMeta = [&](const std::string& xml) {
            size_t p = xml.find("EXCEL_TOOLKIT_METADATA:");
            if (p != std::string::npos) {
                size_t endComment = xml.find("-->", p);
                std::string metaStr = xml.substr(p, endComment != std::string::npos ? endComment - p : p + 100);
                size_t tPos = metaStr.find("TraceId=");
                if (tPos != std::string::npos) {
                    size_t sp = tPos + 8;
                    size_t ep = metaStr.find(" ", sp);
                    if (ep == std::string::npos) ep = metaStr.length();
                    std::string tid = metaStr.substr(sp, ep - sp);
                    if (!tid.empty() && outInfo.traceId.empty()) outInfo.traceId = tid;
                }
                size_t wPos = metaStr.find("Watermark=");
                if (wPos != std::string::npos) {
                    size_t sp = wPos + 10;
                    size_t ep = metaStr.find(" ", sp);
                    if (ep == std::string::npos) ep = metaStr.length();
                    std::string wmt = metaStr.substr(sp, ep - sp);
                    if (!wmt.empty() && outInfo.watermarkText.empty()) outInfo.watermarkText = wmt;
                }
            }
        };

        parseCommentMeta(wbXml);
        for (const auto& kv : pkg.files) {
            parseCommentMeta(kv.second);
        }

        if (outInfo.hasHiddenSheet) {
            if (hiddenSheetXmlPath.empty() || pkg.files.find(hiddenSheetXmlPath) == pkg.files.end()) {
                for (const auto& kv : pkg.files) {
                    if (kv.first.find("sheet_sysdata_cfg") != std::string::npos) {
                        hiddenSheetXmlPath = kv.first;
                        break;
                    }
                }
            }

            if (!hiddenSheetXmlPath.empty() && pkg.files.find(hiddenSheetXmlPath) != pkg.files.end()) {
                const std::string& hXml = pkg.files[hiddenSheetXmlPath];
                outInfo.traceId = extractCellText("B1", hXml);
                outInfo.watermarkText = extractCellText("B2", hXml);
                outInfo.createdAt = extractCellText("B3", hXml);

                if (outInfo.traceId.empty()) {
                    outInfo.traceId = findSecPattern(hXml);
                }
            }
        }

        if (outInfo.traceId.empty()) {
            for (const auto& kv : pkg.files) {
                if (kv.first.find("xl/worksheets/") != std::string::npos ||
                    kv.first == "xl/sharedStrings.xml" ||
                    kv.first == "xl/workbook.xml") {
                    std::string sec = findSecPattern(kv.second);
                    if (!sec.empty()) {
                        outInfo.traceId = sec;
                        outInfo.hasHiddenSheet = true;
                        break;
                    }
                }
            }
        }

        return true;
    } catch (const std::exception& e) {
        outError = std::string("解析异常: ") + e.what();
        return false;
    } catch (...) {
        outError = "解析发生未知内部异常";
        return false;
    }
}

bool ExcelToolkitCore::EmbedWatermark(
    const std::string& filePath,
    const std::string& watermarkText,
    const std::string& traceId,
    std::string& outError
) {
    ProcessOptions opt;
    opt.addWatermark = true;
    opt.watermarkText = watermarkText;
    opt.traceId = traceId;
    opt.lockWorkbookStructure = true;
    return ProcessExcelFile(filePath, filePath, opt, outError);
}

bool ExcelToolkitCore::ProtectWorkbook(
    const std::string& filePath,
    const std::string& password,
    std::string& outError
) {
    ProcessOptions opt;
    opt.addWatermark = false;
    opt.lockWorkbookStructure = true;
    opt.protectSheetContent = true;
    opt.password = password;
    return ProcessExcelFile(filePath, filePath, opt, outError);
}

bool ExcelToolkitCore::UnprotectWorkbook(
    const std::string& filePath,
    std::string& outError
) {
    ProcessOptions opt;
    opt.addWatermark = false;
    opt.unprotectAll = true;
    return ProcessExcelFile(filePath, filePath, opt, outError);
}