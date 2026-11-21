// ==UserScript==
// @name         网页长文/资料纯净提取与一键导出工具 (CleanReader & Exporter)
// @namespace    https://github.com/clean-reader-exporter
// @version      1.0.0
// @description  智能提取网页已渲染的正文内容（文字、图片、表格、公式），过滤广告与导航栏，支持纯净阅读模式、一键导出 PDF / Word 及调用系统高清打印。
// @author       CleanReader
// @match        *://*/*
// @grant        GM_xmlhttpRequest
// @grant        GM_addStyle
// @connect      *
// @require      https://cdnjs.cloudflare.com/ajax/libs/html2pdf.js/0.10.1/html2pdf.bundle.min.js
// @run-at       document-idle
// ==/UserScript==

(function () {
    'use strict';

    const APP_PREFIX = 'cre-exporter';

    // ==================== 0. 跨源/同源 iframe 文档预览框提取子模块 ====================
    // 4 级自适应提取引擎：
    // Tier 1: 递归穿透嵌套子 iframe
    // Tier 2: 整页文档图片 / CSS background-image 页面 / Canvas 画布页 / SVG 页面直接提取
    // Tier 3: 绝对定位排版文档（pdf2htmlEX / 永中DCS / Office HTML 等）通过 Range 实时坐标重排为正常文档流
    // Tier 4: 标准流式 HTML 清洗提取
    async function extractRichHtmlFromDocument(targetWin, targetDoc, canvasCacheMap, isDeep = false) {
        const root = targetDoc && targetDoc.body;
        if (!root) return '';

        const escapeTxt = (s) => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');

        // 辅助：在目标文档同源上下文内将图片 URL 转为 Base64
        async function fetchImgAsDataUrl(url) {
            if (!url) return '';
            if (url.startsWith('data:image/')) return url;
            try {
                const absUrl = new URL(url, targetDoc.baseURI).href;
                const resp = await targetWin.fetch(absUrl, { credentials: 'include' });
                if (resp.ok) {
                    const blob = await resp.blob();
                    if (blob.type && blob.type.startsWith('image/')) {
                        const b64 = await new Promise((res) => {
                            const r = new FileReader();
                            r.onloadend = () => res(typeof r.result === 'string' && r.result.startsWith('data:image/') ? r.result : absUrl);
                            r.onerror = () => res(absUrl);
                            r.readAsDataURL(blob);
                        });
                        return b64;
                    }
                }
                return absUrl;
            } catch (_) {
                try {
                    return new URL(url, targetDoc.baseURI).href;
                } catch (__) {
                    return url;
                }
            }
        }

        // Tier 1: 检查是否存在大尺寸嵌套子 iframe（iframe 套 iframe）
        const innerIframes = Array.from(root.querySelectorAll('iframe')).filter((f) => {
            const r = f.getBoundingClientRect();
            return r.width >= 300 && r.height >= 200;
        });
        if (innerIframes.length > 0) {
            const nestedResults = await Promise.all(innerIframes.map((subFrame) => new Promise((resolve) => {
                try {
                    const subDoc = subFrame.contentDocument || (subFrame.contentWindow && subFrame.contentWindow.document);
                    if (subDoc && subDoc.body) {
                        extractRichHtmlFromDocument(subFrame.contentWindow, subDoc, new Map(), isDeep).then(resolve);
                        return;
                    }
                } catch (_) {}

                if (!subFrame.contentWindow) { resolve(''); return; }
                const reqId = `cre_nested_${Date.now()}_${Math.random().toString(36).slice(2, 8)}`;
                let done = false;
                function onMsg(ev) {
                    if (!ev.data || ev.data.type !== 'CRE_EXTRACT_IFRAME_RES' || ev.data.requestId !== reqId) return;
                    done = true;
                    clearTimeout(tm);
                    targetWin.removeEventListener('message', onMsg);
                    resolve(ev.data.html || '');
                }
                const tm = setTimeout(() => {
                    if (!done) { done = true; targetWin.removeEventListener('message', onMsg); resolve(''); }
                }, 1800);
                targetWin.addEventListener('message', onMsg);
                try {
                    subFrame.contentWindow.postMessage({ type: 'CRE_EXTRACT_IFRAME_REQ', requestId: reqId, deep: isDeep }, '*');
                } catch (_) {
                    clearTimeout(tm); targetWin.removeEventListener('message', onMsg); resolve('');
                }
            })));
            const joinedNested = nestedResults.filter(Boolean).join('\n');
            if (joinedNested.trim().length > 20) {
                return joinedNested;
            }
        }

        // Tier 2: 检查整页文档图片（<img> / CSS background-image 页面 / <canvas> 页面 / 大尺寸 <svg> 页面）
        const visualPages = [];
        const scrollTopOffset = targetWin.scrollY || targetDoc.documentElement.scrollTop || root.scrollTop || 0;

        // 2a. 大尺寸 Canvas 页面
        Array.from(root.querySelectorAll('canvas')).forEach((c) => {
            const r = c.getBoundingClientRect();
            if (r.width < 300 || r.height < 240) return;
            let dataUrl = null;
            try {
                if (c.width >= 300 && c.height >= 240) {
                    const d = c.toDataURL('image/png');
                    if (d && d.length > 2200) dataUrl = d;
                }
            } catch (_) {}
            dataUrl = dataUrl || (canvasCacheMap && canvasCacheMap.get(c));
            if (dataUrl) {
                visualPages.push({ type: 'img', top: scrollTopOffset + r.top, src: dataUrl });
            }
        });

        // 2b. 大尺寸 <img> 文档页
        Array.from(root.querySelectorAll('img')).forEach((img) => {
            const r = img.getBoundingClientRect();
            const rawSrc = img.getAttribute('data-src') || img.getAttribute('data-original') || img.getAttribute('data-actualsrc') || img.getAttribute('src') || img.src;
            if (!rawSrc) return;
            if ((r.width >= 320 && r.height >= 260) || (img.naturalWidth >= 420 && img.naturalHeight >= 420)) {
                visualPages.push({ type: 'img', top: scrollTopOffset + r.top, src: rawSrc });
            }
        });

        // 2c. 大尺寸 CSS background-image 文档页（如原创力/淘豆/Office预览页 div 背景图）
        Array.from(root.querySelectorAll('div, section, li, span')).forEach((el) => {
            const r = el.getBoundingClientRect();
            if (r.width < 320 || r.height < 260) return;
            const bg = targetWin.getComputedStyle(el).backgroundImage;
            if (bg && bg.startsWith('url(')) {
                const m = bg.match(/url\(["']?([^"')]+)["']?\)/);
                if (m && m[1] && !m[1].startsWith('data:image/svg+xml;base64,PHN2ZyB4bWxucz0iaHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmci')) {
                    visualPages.push({ type: 'img', top: scrollTopOffset + r.top, src: m[1] });
                }
            }
        });

        // 2d. 大尺寸 <svg> 文档页（保留内联样式）
        Array.from(root.querySelectorAll('svg')).forEach((svg) => {
            const r = svg.getBoundingClientRect();
            if (r.width >= 320 && r.height >= 260) {
                visualPages.push({ type: 'svg', top: scrollTopOffset + r.top, html: svg.outerHTML });
            }
        });

        if (visualPages.length > 0) {
            visualPages.sort((a, b) => a.top - b.top);
            // 去重相同 URL
            const seen = new Set();
            const uniquePages = visualPages.filter((p) => {
                const key = p.type === 'img' ? p.src : `svg_${Math.round(p.top)}`;
                if (seen.has(key)) return false;
                seen.add(key);
                return true;
            });

            const renderedPages = await Promise.all(uniquePages.map(async (p, idx) => {
                if (p.type === 'svg') {
                    return `<div class="${APP_PREFIX}-canvas-page-item" style="margin:0 auto 16px auto;page-break-inside:avoid;">${p.html}</div>`;
                }
                const finalSrc = await fetchImgAsDataUrl(p.src);
                return `<div class="${APP_PREFIX}-canvas-page-item" style="margin:0 auto 16px auto;page-break-inside:avoid;"><img src="${finalSrc}" alt="文档第 ${idx + 1} 页" class="${APP_PREFIX}-block-img" style="display:block;width:100%;max-width:100%;height:auto;border:1px solid #e2e8f0;" /></div>`;
            }));
            return renderedPages.join('\n');
        }

        // Tier 3: 检测是否为绝对定位排版文档（如 pdf2htmlEX / 永中DCS / WPS HTML 预览框）
        // 这类文档每个文字片段都使用 position:absolute 定位，若直接克隆并删掉 <style> 会导致所有文字错位/重叠或仅剩一个碎片（如 "C.sheeps;cow"）！
        let absoluteTextElCount = 0;
        const sampleEls = root.querySelectorAll('div, span, p');
        for (let i = 0; i < Math.min(sampleEls.length, 300); i++) {
            const el = sampleEls[i];
            if (el.childNodes.length > 0 && (el.textContent || '').trim().length > 0) {
                const st = targetWin.getComputedStyle(el);
                if (st.position === 'absolute') {
                    absoluteTextElCount++;
                }
            }
        }

        if (absoluteTextElCount >= 6) {
            // 使用 TreeWalker + Range.getBoundingClientRect() 获取每一个真实文本片段的视觉坐标 (docTop, left)，按行聚类重排为纯净段落流！
            const items = [];
            const walker = targetDoc.createTreeWalker(root, NodeFilter.SHOW_TEXT, null);
            let textNode;
            while ((textNode = walker.nextNode())) {
                const rawText = textNode.nodeValue;
                if (!rawText || !rawText.trim()) continue;
                const parentEl = textNode.parentElement;
                if (!parentEl) continue;
                const pTag = parentEl.tagName.toUpperCase();
                if (['SCRIPT', 'STYLE', 'NOSCRIPT', 'BUTTON', 'INPUT', 'SELECT', 'TEXTAREA'].includes(pTag)) continue;
                if (!isDeep) {
                    const st = targetWin.getComputedStyle(parentEl);
                    if (st.display === 'none' || st.visibility === 'hidden') continue;
                }
                try {
                    const range = targetDoc.createRange();
                    range.selectNodeContents(textNode);
                    const rect = range.getBoundingClientRect();
                    if (rect.width > 0 || rect.height > 0 || isDeep) {
                        items.push({
                            kind: 'text',
                            top: scrollTopOffset + rect.top,
                            left: rect.left,
                            text: rawText.replace(/\s+/g, ' ')
                        });
                    }
                } catch (_) {}
            }

            // 同时收集内联/独立图片及其坐标
            const inlineImgs = Array.from(root.querySelectorAll('img'));
            for (const img of inlineImgs) {
                const r = img.getBoundingClientRect();
                const rawSrc = img.getAttribute('data-src') || img.getAttribute('data-original') || img.getAttribute('src') || img.src;
                if (!rawSrc || (r.width < 8 && r.height < 8 && !isDeep)) continue;
                const finalSrc = await fetchImgAsDataUrl(rawSrc);
                items.push({
                    kind: 'img',
                    top: scrollTopOffset + r.top,
                    left: r.left,
                    width: r.width,
                    height: r.height,
                    src: finalSrc
                });
            }

            if (items.length > 0) {
                // 按垂直坐标 top 排序，将同一视觉水平基线（差距 <= 9px）内的片段聚成同一行，行内按 left 从左到右拼接
                items.sort((a, b) => (Math.abs(a.top - b.top) <= 9 ? a.left - b.left : a.top - b.top));

                const lines = [];
                items.forEach((item) => {
                    const lastLine = lines[lines.length - 1];
                    if (lastLine && Math.abs(item.top - lastLine.top) <= 9) {
                        lastLine.parts.push(item);
                    } else {
                        lines.push({ top: item.top, parts: [item] });
                    }
                });

                const htmlLines = lines.map((line) => {
                    line.parts.sort((a, b) => a.left - b.left);
                    let lineHtml = '';
                    line.parts.forEach((p, idx) => {
                        if (p.kind === 'img') {
                            const isBig = p.width >= 200 && p.height >= 24;
                            const st = isBig
                                ? 'max-width:100%;height:auto;display:block;margin:8px 0;'
                                : 'max-width:100%;height:auto;display:inline-block;vertical-align:middle;margin:0 2px;';
                            lineHtml += `<img src="${p.src}" style="${st}" />`;
                        } else {
                            const prev = idx > 0 ? line.parts[idx - 1] : null;
                            const gap = prev ? p.left - (prev.left + (prev.text ? prev.text.length * 6 : 20)) : 0;
                            const spacer = prev && gap > 18 ? '&nbsp;&nbsp;&nbsp;&nbsp;' : (lineHtml && !lineHtml.endsWith(' ') && !p.text.startsWith(' ') ? ' ' : '');
                            lineHtml += spacer + escapeTxt(p.text);
                        }
                    });
                    return `<p style="margin:0 0 8px 0;line-height:1.75;text-align:left;">${lineHtml.trim()}</p>`;
                });

                return htmlLines.join('\n');
            }
        }

        // Tier 4: 标准流式 HTML 文档清洗提取（不误删 <svg> 内部样式，移除 position:absolute/fixed 错位）
        const clone = root.cloneNode(true);
        const origCanvases = root.querySelectorAll('canvas');
        const cloneCanvases = clone.querySelectorAll('canvas');
        if (origCanvases.length === cloneCanvases.length) {
            origCanvases.forEach((origC, idx) => {
                let dataUrl = null;
                try {
                    if (origC.width >= 24 && origC.height >= 18) {
                        const d = origC.toDataURL('image/png');
                        if (d && d.length > 1400) dataUrl = d;
                    }
                } catch (_) {}
                dataUrl = dataUrl || (canvasCacheMap && canvasCacheMap.get(origC));
                if (dataUrl) {
                    const img = targetDoc.createElement('img');
                    img.src = dataUrl;
                    img.className = `${APP_PREFIX}-block-img`;
                    img.style.display = 'block';
                    img.style.maxWidth = '100%';
                    img.style.height = 'auto';
                    cloneCanvases[idx].replaceWith(img);
                } else {
                    cloneCanvases[idx].remove();
                }
            });
        }

        if (!isDeep) {
            const origEls = root.querySelectorAll('*');
            const cloneEls = clone.querySelectorAll('*');
            if (origEls.length === cloneEls.length) {
                for (let i = origEls.length - 1; i >= 0; i--) {
                    const tag = origEls[i].tagName.toUpperCase();
                    if (['BR', 'IMG', 'SVG', 'MATH'].includes(tag)) continue;
                    const st = targetWin.getComputedStyle(origEls[i]);
                    if (st.display === 'none' || st.visibility === 'hidden' || parseFloat(st.opacity) === 0) {
                        cloneEls[i].remove();
                    }
                }
            }
        }

        clone.querySelectorAll('script, noscript, iframe, button, input, select, textarea, nav, aside, footer').forEach(n => n.remove());
        clone.querySelectorAll('style').forEach(st => {
            if (!st.closest('svg')) st.remove();
        });

        const imgs = Array.from(clone.querySelectorAll('img'));
        await Promise.all(imgs.map(async (img) => {
            const lazy = img.getAttribute('data-src') || img.getAttribute('data-original') || img.getAttribute('data-actualsrc');
            if ((!img.getAttribute('src') || img.getAttribute('src').startsWith('data:image/svg')) && lazy) {
                img.setAttribute('src', lazy);
            }
            const rawSrc = img.getAttribute('src') || img.src;
            if (!rawSrc) return;
            const finalSrc = await fetchImgAsDataUrl(rawSrc);
            img.setAttribute('src', finalSrc);
            img.style.maxWidth = '100%';
            img.style.height = 'auto';
        }));

        clone.querySelectorAll('*').forEach((el) => {
            if (el.style) {
                if (['fixed', 'sticky', 'absolute'].includes(el.style.position)) {
                    el.style.position = 'static';
                }
                el.style.maxHeight = 'none';
                el.style.overflow = 'visible';
                if (el.style.height && el.style.height.endsWith('px') && ['DIV', 'SECTION', 'MAIN', 'ARTICLE'].includes(el.tagName)) {
                    el.style.height = 'auto';
                }
            }
        });

        return clone.innerHTML.trim();
    }

    if (window.top !== window.self) {
        const iframeCanvasCache = new Map();

        function cacheIframeCanvases() {
            document.querySelectorAll('canvas').forEach((c) => {
                try {
                    if (c.width >= 30 && c.height >= 20) {
                        const dataUrl = c.toDataURL('image/png');
                        if (dataUrl && dataUrl.length > 1400) {
                            const prev = iframeCanvasCache.get(c);
                            if (!prev || dataUrl.length >= prev.length) {
                                iframeCanvasCache.set(c, dataUrl);
                            }
                        }
                    }
                } catch (_) {}
            });
        }

        let iframeScrollTimer = null;
        window.addEventListener('scroll', () => {
            clearTimeout(iframeScrollTimer);
            iframeScrollTimer = setTimeout(cacheIframeCanvases, 250);
        }, { passive: true });
        setTimeout(cacheIframeCanvases, 800);

        window.addEventListener('message', async (event) => {
            const data = event.data;
            if (!data || data.type !== 'CRE_EXTRACT_IFRAME_REQ' || !data.requestId) return;
            try {
                cacheIframeCanvases();
                const html = await extractRichHtmlFromDocument(window, document, iframeCanvasCache, Boolean(data.deep));
                if (event.source && typeof event.source.postMessage === 'function') {
                    event.source.postMessage({
                        type: 'CRE_EXTRACT_IFRAME_RES',
                        requestId: data.requestId,
                        html
                    }, '*');
                }
            } catch (_) {}
        });

        return;
    }

    // ==================== 1. 样式注入 ====================
    const CSS_STYLES = `
        /* 悬浮按钮与面板 */
        #${APP_PREFIX}-fab-container {
            position: fixed;
            right: 24px;
            bottom: 28px;
            z-index: 2147483646;
            font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, "Noto Sans SC", sans-serif;
            user-select: none;
        }

        #${APP_PREFIX}-fab-btn {
            display: flex;
            align-items: center;
            gap: 8px;
            background: linear-gradient(135deg, #2563eb, #1d4ed8);
            color: #ffffff;
            border: none;
            border-radius: 999px;
            padding: 10px 18px;
            font-size: 14px;
            font-weight: 600;
            cursor: pointer;
            box-shadow: 0 8px 20px rgba(37, 99, 235, 0.35);
            transition: transform 0.2s ease, box-shadow 0.2s ease, background 0.2s ease;
        }

        #${APP_PREFIX}-fab-btn:hover {
            transform: translateY(-2px);
            box-shadow: 0 12px 24px rgba(37, 99, 235, 0.45);
            background: linear-gradient(135deg, #1d4ed8, #1e40af);
        }

        #${APP_PREFIX}-panel {
            position: absolute;
            right: 0;
            bottom: 54px;
            width: 230px;
            background: #ffffff;
            border: 1px solid #e5e7eb;
            border-radius: 14px;
            box-shadow: 0 16px 36px rgba(15, 23, 42, 0.16);
            padding: 10px;
            display: none;
            flex-direction: column;
            gap: 6px;
        }

        #${APP_PREFIX}-panel.open {
            display: flex;
        }

        .${APP_PREFIX}-panel-title {
            font-size: 12px;
            font-weight: 600;
            color: #6b7280;
            padding: 4px 8px 6px;
            border-bottom: 1px solid #f3f4f6;
            margin-bottom: 2px;
            display: flex;
            justify-content: space-between;
            align-items: center;
        }

        .${APP_PREFIX}-menu-item {
            display: flex;
            align-items: center;
            gap: 10px;
            width: 100%;
            text-align: left;
            background: transparent;
            border: none;
            border-radius: 8px;
            padding: 9px 12px;
            font-size: 13.5px;
            color: #1f2937;
            font-weight: 500;
            cursor: pointer;
            transition: background 0.15s ease, color 0.15s ease;
        }

        .${APP_PREFIX}-menu-item:hover {
            background: #eff6ff;
            color: #1d4ed8;
        }

        .${APP_PREFIX}-menu-item:disabled {
            opacity: 0.55;
            cursor: not-allowed;
        }

        /* 手动点选高亮框 */
        .${APP_PREFIX}-picker-highlight {
            outline: 3px dashed #2563eb !important;
            background-color: rgba(37, 99, 235, 0.06) !important;
            cursor: crosshair !important;
        }

        /* 纯净阅读模式遮罩层 */
        #${APP_PREFIX}-clean-overlay {
            position: fixed;
            inset: 0;
            background: #f1f5f9;
            z-index: 2147483647;
            display: none;
            flex-direction: column;
            overflow: hidden;
            font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "PingFang SC", "Microsoft YaHei", sans-serif;
        }

        #${APP_PREFIX}-clean-overlay.active {
            display: flex;
        }

        #${APP_PREFIX}-toolbar {
            display: flex;
            align-items: center;
            justify-content: space-between;
            padding: 12px 24px;
            background: #ffffff;
            border-bottom: 1px solid #e2e8f0;
            box-shadow: 0 1px 3px rgba(0, 0, 0, 0.05);
            flex-shrink: 0;
        }

        .${APP_PREFIX}-toolbar-left {
            display: flex;
            align-items: center;
            gap: 12px;
        }

        .${APP_PREFIX}-toolbar-title {
            font-size: 15px;
            font-weight: 700;
            color: #0f172a;
        }

        .${APP_PREFIX}-toolbar-badge {
            font-size: 12px;
            padding: 2px 8px;
            border-radius: 999px;
            background: #dbeafe;
            color: #1e40af;
            font-weight: 500;
        }

        .${APP_PREFIX}-toolbar-actions {
            display: flex;
            align-items: center;
            gap: 10px;
        }

        .${APP_PREFIX}-btn {
            display: inline-flex;
            align-items: center;
            gap: 6px;
            padding: 7px 14px;
            border-radius: 8px;
            font-size: 13px;
            font-weight: 600;
            cursor: pointer;
            border: 1px solid #cbd5e1;
            background: #ffffff;
            color: #334155;
            transition: all 0.15s ease;
        }

        .${APP_PREFIX}-btn:hover {
            background: #f8fafc;
            border-color: #94a3b8;
        }

        .${APP_PREFIX}-btn-primary {
            background: #2563eb;
            color: #ffffff;
            border-color: #2563eb;
        }

        .${APP_PREFIX}-btn-primary:hover {
            background: #1d4ed8;
            border-color: #1d4ed8;
        }

        .${APP_PREFIX}-btn-danger {
            color: #dc2626;
            border-color: #fecaca;
            background: #fef2f2;
        }

        .${APP_PREFIX}-btn-danger:hover {
            background: #fee2e2;
        }

        #${APP_PREFIX}-scroll-area {
            flex: 1;
            overflow-y: auto;
            padding: 32px 16px 64px;
        }

        /* 纯净阅读纸张容器与 PDF 导出独立容器共用排版样式 */
        #${APP_PREFIX}-paper,
        .${APP_PREFIX}-pdf-render-box {
            max-width: 820px;
            margin: 0 auto;
            background: #ffffff;
            color: #111827;
            padding: 48px 56px;
            border-radius: 8px;
            box-shadow: 0 4px 20px rgba(15, 23, 42, 0.08);
            line-height: 1.75;
            font-size: 16px;
            word-wrap: break-word;
            overflow-wrap: break-word;
        }

        .${APP_PREFIX}-pdf-render-box {
            width: 100% !important;
            max-width: 100% !important;
            box-sizing: border-box !important;
            margin: 0 !important;
            padding: 4px 10px !important;
            box-shadow: none !important;
            border-radius: 0 !important;
            background: #ffffff !important;
        }

        .${APP_PREFIX}-pdf-render-box * {
            box-sizing: border-box !important;
        }

        /* 关键修复：html2canvas 在渲染跨行内联元素 (span/code/a) 的背景色时，
           会使用整个跨行包围盒 (getBoundingClientRect) 绘制不透明背景并遮盖同行前序文字。
           因此在 PDF 渲染容器中强制将内联元素背景设为透明，并使用左对齐排版。 */
        .${APP_PREFIX}-pdf-render-box span,
        .${APP_PREFIX}-pdf-render-box code,
        .${APP_PREFIX}-pdf-render-box a,
        .${APP_PREFIX}-pdf-render-box em,
        .${APP_PREFIX}-pdf-render-box i,
        .${APP_PREFIX}-pdf-render-box b,
        .${APP_PREFIX}-pdf-render-box strong {
            background: transparent !important;
            background-color: transparent !important;
        }

        #${APP_PREFIX}-paper[contenteditable="true"] {
            outline: 2px dashed #3b82f6;
            outline-offset: 4px;
        }

        #${APP_PREFIX}-paper h1,
        #${APP_PREFIX}-paper h2,
        #${APP_PREFIX}-paper h3,
        #${APP_PREFIX}-paper h4,
        .${APP_PREFIX}-pdf-render-box h1,
        .${APP_PREFIX}-pdf-render-box h2,
        .${APP_PREFIX}-pdf-render-box h3,
        .${APP_PREFIX}-pdf-render-box h4 {
            color: #0f172a;
            line-height: 1.35;
            margin-top: 1.2em;
            margin-bottom: 0.6em;
            font-weight: 700;
            page-break-after: avoid;
            break-after: avoid;
        }

        #${APP_PREFIX}-paper h1,
        .${APP_PREFIX}-pdf-render-box h1 {
            font-size: 24px;
            margin-top: 0;
            padding-bottom: 12px;
            border-bottom: 2px solid #e5e7eb;
        }

        #${APP_PREFIX}-paper h2,
        .${APP_PREFIX}-pdf-render-box h2 { font-size: 20px; }
        #${APP_PREFIX}-paper h3,
        .${APP_PREFIX}-pdf-render-box h3 { font-size: 17px; }

        #${APP_PREFIX}-paper p {
            margin: 0 0 1em 0;
            text-align: justify;
        }

        .${APP_PREFIX}-pdf-render-box p {
            margin: 0 0 1em 0;
            text-align: left !important;
        }

        #${APP_PREFIX}-paper img,
        #${APP_PREFIX}-paper svg,
        #${APP_PREFIX}-paper canvas,
        .${APP_PREFIX}-pdf-render-box img,
        .${APP_PREFIX}-pdf-render-box svg,
        .${APP_PREFIX}-pdf-render-box canvas {
            max-width: 100% !important;
            height: auto !important;
            display: inline-block;
            vertical-align: middle;
            page-break-inside: avoid;
            break-inside: avoid;
        }

        #${APP_PREFIX}-paper img.${APP_PREFIX}-block-img,
        .${APP_PREFIX}-pdf-render-box img.${APP_PREFIX}-block-img,
        #${APP_PREFIX}-paper .${APP_PREFIX}-canvas-page-item img,
        .${APP_PREFIX}-pdf-render-box .${APP_PREFIX}-canvas-page-item img {
            display: block !important;
            margin: 14px auto !important;
        }

        #${APP_PREFIX}-paper table,
        .${APP_PREFIX}-pdf-render-box table {
            width: 100%;
            border-collapse: collapse;
            margin: 16px 0;
            font-size: 14px;
            page-break-inside: avoid;
            break-inside: avoid;
        }

        #${APP_PREFIX}-paper th,
        #${APP_PREFIX}-paper td,
        .${APP_PREFIX}-pdf-render-box th,
        .${APP_PREFIX}-pdf-render-box td {
            border: 1px solid #cbd5e1;
            padding: 8px 12px;
            text-align: left;
            vertical-align: top;
        }

        #${APP_PREFIX}-paper th,
        .${APP_PREFIX}-pdf-render-box th {
            background-color: #f8fafc;
            font-weight: 600;
        }

        #${APP_PREFIX}-paper table.${APP_PREFIX}-layout-table,
        #${APP_PREFIX}-paper table.${APP_PREFIX}-layout-table tr,
        #${APP_PREFIX}-paper table.${APP_PREFIX}-layout-table td,
        #${APP_PREFIX}-paper table.${APP_PREFIX}-layout-table th,
        .${APP_PREFIX}-pdf-render-box table.${APP_PREFIX}-layout-table,
        .${APP_PREFIX}-pdf-render-box table.${APP_PREFIX}-layout-table tr,
        .${APP_PREFIX}-pdf-render-box table.${APP_PREFIX}-layout-table td,
        .${APP_PREFIX}-pdf-render-box table.${APP_PREFIX}-layout-table th {
            border: none !important;
            background: transparent !important;
            padding: 4px 12px 4px 0 !important;
            margin: 4px 0 !important;
        }

        #${APP_PREFIX}-paper pre,
        #${APP_PREFIX}-paper blockquote,
        .${APP_PREFIX}-pdf-render-box pre,
        .${APP_PREFIX}-pdf-render-box blockquote {
            background: #f8fafc;
            border-left: 4px solid #3b82f6;
            padding: 12px 16px;
            margin: 14px 0;
            border-radius: 4px;
            overflow-x: auto;
            page-break-inside: avoid;
            break-inside: avoid;
        }

        /* Toast 提示 */
        #${APP_PREFIX}-toast {
            position: fixed;
            top: 20px;
            left: 50%;
            transform: translateX(-50%);
            background: rgba(15, 23, 42, 0.92);
            color: #ffffff;
            padding: 10px 20px;
            border-radius: 999px;
            font-size: 13.5px;
            font-weight: 500;
            z-index: 2147483647;
            box-shadow: 0 8px 24px rgba(0, 0, 0, 0.2);
            display: none;
            pointer-events: none;
        }

        /* 打印专用样式：无论是点击“系统打印”还是在纯净阅读模式下直接按 Ctrl+P，均完整打印所有分页 */
        @media print {
            @page {
                size: A4 portrait;
                margin: 15mm 15mm 15mm 15mm;
            }

            html:has(body.${APP_PREFIX}-printing),
            html:has(body.${APP_PREFIX}-clean-open),
            body.${APP_PREFIX}-printing,
            body.${APP_PREFIX}-clean-open {
                background: #ffffff !important;
                margin: 0 !important;
                padding: 0 !important;
                height: auto !important;
                min-height: 0 !important;
                max-height: none !important;
                overflow: visible !important;
            }

            body.${APP_PREFIX}-printing > *:not(#${APP_PREFIX}-clean-overlay),
            body.${APP_PREFIX}-clean-open > *:not(#${APP_PREFIX}-clean-overlay) {
                display: none !important;
            }

            body.${APP_PREFIX}-printing #${APP_PREFIX}-clean-overlay,
            body.${APP_PREFIX}-clean-open #${APP_PREFIX}-clean-overlay {
                position: static !important;
                display: block !important;
                width: 100% !important;
                height: auto !important;
                max-height: none !important;
                background: #ffffff !important;
                overflow: visible !important;
            }

            body.${APP_PREFIX}-printing #${APP_PREFIX}-toolbar,
            body.${APP_PREFIX}-clean-open #${APP_PREFIX}-toolbar,
            body.${APP_PREFIX}-printing #${APP_PREFIX}-fab-container,
            body.${APP_PREFIX}-clean-open #${APP_PREFIX}-fab-container,
            body.${APP_PREFIX}-printing #${APP_PREFIX}-toast,
            body.${APP_PREFIX}-clean-open #${APP_PREFIX}-toast {
                display: none !important;
            }

            body.${APP_PREFIX}-printing #${APP_PREFIX}-scroll-area,
            body.${APP_PREFIX}-clean-open #${APP_PREFIX}-scroll-area {
                padding: 0 !important;
                margin: 0 !important;
                height: auto !important;
                max-height: none !important;
                overflow: visible !important;
            }

            body.${APP_PREFIX}-printing #${APP_PREFIX}-paper,
            body.${APP_PREFIX}-clean-open #${APP_PREFIX}-paper {
                max-width: 100% !important;
                width: 100% !important;
                padding: 0 !important;
                margin: 0 !important;
                box-shadow: none !important;
                border: none !important;
                height: auto !important;
                max-height: none !important;
                overflow: visible !important;
            }
        }
    `;

    function injectStyles() {
        if (typeof GM_addStyle === 'function') {
            GM_addStyle(CSS_STYLES);
        } else {
            const styleEl = document.createElement('style');
            styleEl.id = `${APP_PREFIX}-styles`;
            styleEl.textContent = CSS_STYLES;
            document.head.appendChild(styleEl);
        }
    }

    // ==================== 2. 智能正文识别与 DOM 清洗 ====================

    // 常见的噪音节点选择器（已移除过于宽泛的 knowledge / collection / related / vip 通配符，防止误删组卷网/学科网的试题容器）
    const NOISE_SELECTORS = [
        'script', 'style', 'noscript', 'iframe', 'nav', 'aside', 'footer',
        'button', 'input', 'select', 'textarea',
        '[role="navigation"]', '[role="banner"]', '[role="complementary"]', '[role="contentinfo"]',
        '.sidebar', '.side-bar', '.sider', '.left-sidebar', '.right-sidebar',
        '.nav', '.navbar', '.navigation', '.site-header', '.site-footer', '.footer',
        '.ad', '.ads', '.advert', '.advertisement', '.banner-ad', '.sponsor', '.recommend',
        '.related-posts', '.related-docs', '.comments', '#comments', '.comment-list',
        '.share', '.social-share', '.toolbar', '.tool-bar', '.action-bar', '.bottom-bar',
        '.breadcrumb', '.breadcrumbs', '.pagination', '.floating',
        '.quest-menu', '.ctrl-box', '.oper-box', '.basket-box',
        '[class*="sidebar"]', '[class*="side-bar"]', '[class*="ai-assistant"]',
        '[class*="copilot"]', '[class*="fold-page"]', '[class*="try-read"]',
        '[class*="search-box"]', '[class*="search-bar"]', '[class*="next-doc"]',
        `#${APP_PREFIX}-fab-container`, `#${APP_PREFIX}-clean-overlay`, `#${APP_PREFIX}-toast`, `#${APP_PREFIX}-print-iframe`
    ];

    // 常见正文、题库试卷与文档阅读器优先候选选择器
    const CANDIDATE_SELECTORS = [
        'article',
        'main',
        '[role="main"]',
        '.paper-cnt',
        '.paper-content',
        '.paper-body',
        '.paper-detail',
        '#paper-body',
        '.quest-list',
        '.question-list',
        '.tk-quest-list',
        '.question-container',
        '#question-list',
        '.sec-list',
        '.exam-cnt',
        '.exam-content',
        '.exam-body',
        '.exam-detail',
        '.article-content',
        '.post-content',
        '.entry-content',
        '.doc-content',
        '.reader-container',
        '#reader-container',
        '.doc-reader',
        '.reader-wrap',
        '.page-container',
        '#page-container',
        '.markdown-body',
        '.rich_media_content',
        '#article-content'
    ];

    const NOISE_CLASS_REGEX = /(comment|sidebar|side-bar|sider|site-footer|navbar|banner-ad|advertisement|social-share|search-box|search-bar|breadcrumb|ai-assistant|copilot|basket-box)/i;

    /**
     * 判断是否为本脚本自己注入的 UI 节点（防止二次提取时把脚本弹窗自己识别进去）
     */
    function isOwnPluginElement(el) {
        if (!el || el.nodeType !== Node.ELEMENT_NODE) return false;
        if (el === document.body || el === document.documentElement) return false;
        return Boolean(
            typeof el.closest === 'function' &&
            el.closest(`#${APP_PREFIX}-fab-container, #${APP_PREFIX}-clean-overlay, #${APP_PREFIX}-toast, #${APP_PREFIX}-print-iframe`)
        );
    }

    /**
     * 判断节点在当前页面上是否对用户真正可见
     * （递归检查所有祖先节点的 opacity:0 / display:none / visibility:hidden / 0高度裁剪 / 屏幕外弹窗，
     *   防止将组卷网隐藏的“试卷分析/细目表”弹窗或 ECharts 隐藏饼图误判为可见正文）
     */
    function isElementVisible(el) {
        if (!el || el.nodeType !== Node.ELEMENT_NODE) return false;
        if (isOwnPluginElement(el)) return false;

        const rect = el.getBoundingClientRect();
        if (rect.width <= 2 || rect.height <= 2) return false;
        // 排除位于屏幕水平可视范围之外的隐藏抽屉或移出屏幕的弹窗
        if (rect.right <= 0 || rect.left >= window.innerWidth) return false;

        let cur = el;
        while (cur && cur !== document.documentElement) {
            const style = window.getComputedStyle(cur);
            if (
                style.display === 'none' ||
                style.visibility === 'hidden' ||
                parseFloat(style.opacity) === 0
            ) {
                return false;
            }
            if (cur !== el) {
                const curRect = cur.getBoundingClientRect();
                const ov = `${style.overflow} ${style.overflowX} ${style.overflowY}`;
                if ((curRect.height <= 2 || curRect.width <= 2) && /hidden|clip/.test(ov)) {
                    return false;
                }
            }
            cur = cur.parentElement;
        }
        return true;
    }

    /**
     * 判断节点是否属于弹窗/模态框或“试卷分析 / 细目表分析”统计面板（而非试卷正文题目）
     */
    function isModalOrAnalysisDialog(el) {
        if (!el || el.nodeType !== Node.ELEMENT_NODE) return false;
        if (el.closest('.el-dialog, .el-drawer, .ant-modal, .ant-drawer, .layui-layer, .modal, .dialog, [role="dialog"], [aria-modal="true"]')) {
            return true;
        }
        const txt = (el.innerText || '').trim();
        const hasChoiceOptions = /[AＡ]\s*[.．、][\s\S]{1,120}[BＢ]\s*[.．、]/.test(txt);
        // 只要文本开头 120 字内出现“试卷分析”且包含“试卷难度 / 细目表分析 / 知识点分析”，且不含 A./B. 题目选项，就一定是统计分析面板，绝不是试卷题目
        if (!hasChoiceOptions && /试卷分析/.test(txt.slice(0, 120)) && /(试卷难度|细目表分析|知识点分析)/.test(txt)) {
            return true;
        }
        // 纯细目表分析表格块（包含“细目表分析”与“难度系数/详细知识点”，不含 A./B. 选项）
        if (!hasChoiceOptions && txt.length < 4000 && /细目表分析/.test(txt.slice(0, 150)) && /详细知识点/.test(txt.slice(0, 250))) {
            return true;
        }
        return false;
    }

    /**
     * 判断节点是否位于桌面端网页的左右边缘侧边栏区域
     * 注意：仅针对真正的大块侧边栏容器（宽 >= 80px 且高 >= 180px 的块级容器），绝不误伤正文内部的选项、公式图片或表格单元格！
     */
    function isEdgeSidebarRegion(el) {
        if (!el || el.nodeType !== Node.ELEMENT_NODE) return false;
        if (window.innerWidth < 920) return false;
        const tag = el.tagName.toUpperCase();
        if (!['DIV', 'ASIDE', 'SECTION', 'NAV', 'UL'].includes(tag)) return false;
        if (el.closest('table, tr, td, p, li, [class*="quest"], [class*="timu"], [class*="exam-item"]')) return false;

        const rect = el.getBoundingClientRect();
        // 必须是具有一定高度与宽度的独立侧边栏块，排除行内文字、公式图片和单行选项
        if (rect.width < 80 || rect.height < 180 || rect.width > window.innerWidth * 0.42) return false;
        if (rect.right < window.innerWidth * 0.22 || rect.left > window.innerWidth * 0.72) {
            return true;
        }
        return false;
    }

    // 以 Canvas DOM 引用本身作为 Key 缓存已绘制的画布图像（彻底避免不同画布因父级 ID 相同而互相覆盖为“1页”）
    const canvasBitmapCache = new Map();

    /**
     * 判断 Canvas 是否为 ECharts / 统计图表画布（非文档正文页）
     */
    function isChartCanvas(canvas) {
        if (!canvas) return false;
        if (canvas.hasAttribute('_echarts_instance_') || canvas.getAttribute('data-zr-dom-id')) return true;
        if (canvas.closest('[_echarts_instance_], [class*="echarts"], [class*="chart"], [id*="chart"]')) return true;
        if (isModalOrAnalysisDialog(canvas)) return true;
        return false;
    }

    /**
     * 检测 Canvas 是否已经实际绘制了可见内容（排除空白占位的懒加载 Canvas 与统计饼图）
     */
    function getDrawnCanvasDataUrl(canvas) {
        if (!canvas || !isElementVisible(canvas) || isChartCanvas(canvas)) return null;
        const rect = canvas.getBoundingClientRect();
        if (rect.width < 24 || rect.height < 18 || canvas.width < 24 || canvas.height < 18) return null;
        try {
            const dataUrl = canvas.toDataURL('image/png');
            if (dataUrl && dataUrl.length > 1400) {
                return dataUrl;
            }
        } catch (_) {
            // 跨域污染的 canvas
        }
        return null;
    }

    /**
     * 在用户滚动浏览时自动以 Canvas 元素引用为 Key 缓存已渲染完毕的画布（防止虚拟滚动清空上方已读画布）
     */
    function scanAndCacheVisibleCanvases() {
        const allCanvases = Array.from(document.querySelectorAll('canvas'));
        allCanvases.forEach((c) => {
            if (isOwnPluginElement(c) || isChartCanvas(c)) return;
            if (c.closest('aside, nav, footer, header, [class*="sidebar"], [class*="ai-assistant"]')) {
                return;
            }
            if (isEdgeSidebarRegion(c)) return;

            const dataUrl = getDrawnCanvasDataUrl(c);
            if (!dataUrl) return;

            const docTop = Math.round(window.scrollY + c.getBoundingClientRect().top);
            const existing = canvasBitmapCache.get(c);
            if (!existing || dataUrl.length >= existing.dataUrl.length) {
                canvasBitmapCache.set(c, {
                    dataUrl,
                    top: docTop,
                    width: c.getBoundingClientRect().width,
                    height: c.getBoundingClientRect().height
                });
            }
        });
    }

    /**
     * 判断 iframe 是否为文档/正文内容预览框（排除广告 iframe 与本脚本打印 iframe）
     */
    function isContentIframe(iframe) {
        if (!iframe || iframe.tagName !== 'IFRAME') return false;
        if (isOwnPluginElement(iframe)) return false;
        const meta = `${iframe.src || ''} ${iframe.id || ''} ${typeof iframe.className === 'string' ? iframe.className : ''}`.toLowerCase();
        if (/(doubleclick|googlesyndication|adservice|cpro|tanx|union|hm\.baidu|pos\.baidu|advert|banner-ad)/i.test(meta)) {
            return false;
        }
        if (!isElementVisible(iframe) || isEdgeSidebarRegion(iframe)) return false;
        const rect = iframe.getBoundingClientRect();
        return rect.width >= 300 && rect.height >= 180;
    }

    /**
     * 从单个文档预览 iframe 中提取纯净 HTML（优先同源直读，跨源时通过 postMessage 与子模块通信提取）
     */
    function extractHtmlFromIframe(iframe, isDeep = false) {
        return new Promise((resolve) => {
            if (!iframe) {
                resolve('');
                return;
            }

            // 1. 尝试同源直接调用 4 级自适应提取引擎读取 contentDocument
            try {
                const doc = iframe.contentDocument || (iframe.contentWindow && iframe.contentWindow.document);
                if (doc && doc.body) {
                    extractRichHtmlFromDocument(iframe.contentWindow, doc, new Map(), isDeep)
                        .then((html) => {
                            if (html && html.trim()) {
                                resolve(html);
                            } else {
                                // 若同源直读为空，仍尝试 postMessage 通道
                                sendPostMessageRequest();
                            }
                        })
                        .catch(() => sendPostMessageRequest());
                    return;
                }
            } catch (_) {
                // 跨源 iframe，转入 postMessage 通道
            }

            function sendPostMessageRequest() {
                if (!iframe.contentWindow) {
                    resolve('');
                    return;
                }
                const requestId = `cre_${Date.now()}_${Math.random().toString(36).slice(2, 8)}`;
                let settled = false;

                function onMessage(event) {
                    const data = event.data;
                    if (!data || data.type !== 'CRE_EXTRACT_IFRAME_RES' || data.requestId !== requestId) return;
                    settled = true;
                    clearTimeout(timer);
                    window.removeEventListener('message', onMessage);
                    resolve(typeof data.html === 'string' ? data.html : '');
                }

                const timer = setTimeout(() => {
                    if (!settled) {
                        settled = true;
                        window.removeEventListener('message', onMessage);
                        resolve('');
                    }
                }, 2600);

                window.addEventListener('message', onMessage);
                try {
                    iframe.contentWindow.postMessage({
                        type: 'CRE_EXTRACT_IFRAME_REQ',
                        requestId,
                        deep: isDeep
                    }, '*');
                } catch (_) {
                    clearTimeout(timer);
                    window.removeEventListener('message', onMessage);
                    resolve('');
                }
            }

            sendPostMessageRequest();
        });
    }

    /**
     * 批量异步提取容器内所有文档预览 iframe 的内容，返回 Map<HTMLIFrameElement, string>
     */
    async function collectContentIframes(rootNode, isDeep = false) {
        const resultMap = new Map();
        if (!rootNode) return resultMap;

        const iframes = rootNode.tagName === 'IFRAME'
            ? (isContentIframe(rootNode) ? [rootNode] : [])
            : Array.from(rootNode.querySelectorAll('iframe')).filter(isContentIframe);

        if (iframes.length === 0) return resultMap;

        await Promise.all(iframes.map(async (frame) => {
            const html = await extractHtmlFromIframe(frame, isDeep);
            if (html && html.trim().length > 0) {
                resultMap.set(frame, html);
            }
        }));
        return resultMap;
    }

    /**
     * 查找页面上所有已渲染的 A4 文档级大画布（仅针对百度文库、PDF.js 等整页大画布阅读器，彻底排除 ECharts 统计饼图）
     */
    function findDocumentPageCanvases(root = document.body) {
        const allCanvases = Array.from(root.querySelectorAll('canvas'));
        return allCanvases.filter(c => {
            if (isOwnPluginElement(c) || isChartCanvas(c)) return false;
            if (c.closest('aside, nav, footer, header, [class*="sidebar"], [class*="ai-assistant"]')) {
                return false;
            }
            if (isEdgeSidebarRegion(c)) return false;
            const rect = c.getBoundingClientRect();
            // 必须是真正的纵向整页文档大画布（宽 >= 380 且高 >= 420，且排除图表）
            if (rect.width < 380 || rect.height < 420) return false;
            return Boolean(getDrawnCanvasDataUrl(c) || canvasBitmapCache.has(c));
        });
    }

    /**
     * 计算多个 DOM 节点的最小公共祖先容器 (Lowest Common Ancestor)
     */
    function getLowestCommonAncestor(nodes) {
        if (!nodes || nodes.length === 0) return null;
        if (nodes.length === 1) return nodes[0].parentElement || nodes[0];

        const getAncestors = (node) => {
            const list = [];
            let cur = node;
            while (cur && cur !== document.documentElement) {
                list.unshift(cur);
                cur = cur.parentElement;
            }
            return list;
        };

        const firstPath = getAncestors(nodes[0]);
        let lca = document.body;
        for (let i = 0; i < firstPath.length; i++) {
            const candidate = firstPath[i];
            if (nodes.every(n => candidate.contains(n))) {
                lca = candidate;
            } else {
                break;
            }
        }
        return lca;
    }

    /**
     * 计算容器节点的“正文价值得分”
     */
    function scoreElement(el) {
        if (!isElementVisible(el)) return -1;

        const tagName = el.tagName.toLowerCase();
        if (['body', 'html', 'nav', 'aside', 'footer', 'header', 'script', 'style', 'button', 'input'].includes(tagName)) {
            return -1;
        }

        const classAndId = `${typeof el.className === 'string' ? el.className : ''} ${el.id || ''}`.toLowerCase();
        if (NOISE_CLASS_REGEX.test(classAndId)) {
            return -100;
        }

        // 排除“试卷分析 / 细目表分析”等弹窗或统计面板
        if (isModalOrAnalysisDialog(el)) {
            return -1000;
        }

        // 桌面端位于左右边缘窄列的容器大幅降权（剔除左侧广告栏、右侧“试题统计/试卷下载”侧边栏等）
        if (isEdgeSidebarRegion(el)) {
            return -120;
        }

        const text = (el.innerText || '').trim();
        const docCanvases = findDocumentPageCanvases(el);
        const contentIframes = Array.from(el.querySelectorAll('iframe')).filter(isContentIframe);

        // 统计容器内的试题图片条或试题配图（兼容组卷网等将每道题或公式渲染为图片 <img> 的题库网站）
        const questionImages = Array.from(el.querySelectorAll('img')).filter(img => {
            const r = img.getBoundingClientRect();
            return r.width >= 200 && r.height >= 22;
        }).length;
        const inlineFormulaImages = Array.from(el.querySelectorAll('img')).filter(img => {
            const r = img.getBoundingClientRect();
            return r.width >= 10 && r.height >= 10;
        }).length;

        // 统计真实试卷特征：选择题选项组 (A. ... B. ...) 与带题号题目 (1. / 2. / 3.)
        const optionGroups = (text.match(/[AＡ]\s*[.．、][\s\S]{1,120}[BＢ]\s*[.．、]/g) || []).length;
        const numberedQuestions = (text.match(/(?:^|\n)\s*\d{1,2}\s*[.．、]/g) || []).length;

        if (text.length < 30 && docCanvases.length === 0 && questionImages === 0 && contentIframes.length === 0) return 0;

        let score = 0;

        // 0. 包含内嵌文档预览 iframe 或独立内部滚动文档框的高权重加分
        score += contentIframes.length * 600;
        const elStyle = window.getComputedStyle(el);
        if ((elStyle.overflowY === 'auto' || elStyle.overflowY === 'scroll') && el.scrollHeight > el.clientHeight + 80 && el.getBoundingClientRect().width >= 420) {
            score += 240;
        }

        // 1. 真实试卷题目选项与编号题干给予最高优先级（彻底压倒“试卷分析/细目表”）
        score += Math.min(optionGroups * 150, 1500);
        score += Math.min(numberedQuestions * 40, 800);

        // 2. 语义标签与题库试卷/阅读器类名加分
        if (tagName === 'article') score += 120;
        if (tagName === 'main') score += 90;
        if (/(paper-cnt|paper-content|paper-body|paper-detail|quest-list|question-list|tk-quest|sec-list|exam-cnt|exam-body|exam-list|article|reader|doc-content|post-content|entry-content|markdown)/i.test(classAndId)) {
            score += 140;
        }

        // 3. 包含“一、单选题 / 二、填空题 / 三、解答题”等标准试卷大题标题的容器加分（需排除细目表）
        const sectionHeadingMatches = text.match(/[一二三四五六七八九十]+、\s*(单选|选择|多选|填空|解答|判断|计算|应用|综合|阅读|完形|写作)/g);
        if (sectionHeadingMatches && !/细目表分析/.test(text.slice(0, 300))) {
            score += Math.min(sectionHeadingMatches.length * 160, 480);
        }

        // 4. 文档级已渲染 Canvas 画布高权重加分
        score += docCanvases.length * 260;

        // 5. 段落、题目块（含组卷网 quest / timu / ques / data-qid）、标点符号与文本长度加分
        const paragraphs = el.querySelectorAll('p').length;
        const questionBlocks = el.querySelectorAll('[class*="quest"], [class*="timu"], [class*="ques"], [class*="exam-item"], [data-qid]').length;
        score += Math.min(paragraphs * 10, 300) + Math.min(questionBlocks * 35, 700);

        const punctuationMatches = text.match(/[，。；！？、,.;!?（）()]/g);
        const punctuationCount = punctuationMatches ? punctuationMatches.length : 0;
        score += Math.min(punctuationCount * 4, 450);

        score += Math.min(Math.floor(text.length / 20), 400);

        // 6. 宽幅试题图片、行内公式图片、数学公式、表格加分
        const tables = el.querySelectorAll('table').length;
        const formulas = el.querySelectorAll('.katex, .MathJax, math, mjx-container').length;
        score += questionImages * 45 + Math.min(inlineFormulaImages * 10, 350) + formulas * 12 + Math.min(tables * 15, 90);

        // 7. 链接密度惩罚（注意：组卷网等题库网站的每道题可能整体被 <a> 包裹以便点击看解析，因此检测到试卷选项/题目块时豁免链接密度惩罚！）
        const isExamPaperLike = optionGroups >= 2 || numberedQuestions >= 3 || questionBlocks >= 3;
        if (!isExamPaperLike && docCanvases.length === 0 && questionImages < 3) {
            const links = el.querySelectorAll('a');
            let linkTextLength = 0;
            links.forEach(a => {
                linkTextLength += (a.innerText || '').trim().length;
            });
            const linkDensity = text.length > 0 ? linkTextLength / text.length : 0;
            if (linkDensity > 0.45) {
                score *= 0.15;
            } else if (linkDensity > 0.28) {
                score *= 0.55;
            }
        }

        // 7. 宽度与视觉居中修正：惩罚包含左右侧边栏的整页超宽外层 Wrapper，奖励居中主列
        const rect = el.getBoundingClientRect();
        if (window.innerWidth >= 1000) {
            if (rect.width >= window.innerWidth * 0.90) {
                score *= 0.55;
            } else if (rect.left >= 40 && rect.right <= window.innerWidth - 180 && rect.width >= 450) {
                score *= 1.3;
            }
        }

        // 8. 内部包含侧边栏/AI助手/试题统计栏等噪音块的外层大容器扣分，促使向内收敛到纯净试卷容器
        const noiseChildrenCount = el.querySelectorAll('aside, nav, .sidebar, [class*="sidebar"], [class*="recommend"], [class*="ai-assistant"]').length;
        if (noiseChildrenCount > 0) {
            score *= Math.max(0.35, 1 - noiseChildrenCount * 0.18);
        }

        return score;
    }

    /**
     * 自动识别页面中的核心正文容器
     */
    function detectMainContentNode() {
        const wasOverlayShown = overlayEl && overlayEl.classList.contains('active') && overlayEl.style.display !== 'none';
        if (wasOverlayShown) {
            overlayEl.style.display = 'none';
        }

        try {
            // 优先检查：如果页面正中央存在大尺寸文档预览 iframe（如在线 Word/PDF 下拉滚动预览框），直接优先定位该 iframe
            const pageContentIframes = Array.from(document.querySelectorAll('iframe')).filter(isContentIframe);
            if (pageContentIframes.length > 0) {
                const largestIframe = pageContentIframes.sort((a, b) => {
                    const ra = a.getBoundingClientRect();
                    const rb = b.getBoundingClientRect();
                    return (rb.width * rb.height) - (ra.width * ra.height);
                })[0];
                const lRect = largestIframe.getBoundingClientRect();
                if (lRect.width >= 420 && lRect.height >= 260) {
                    return largestIframe;
                }
            }

            // 仅当页面不存在题库试题块，且存在 >=2 个纵向 A4 文档级大画布（如百度文库）时，才走画布 LCA 快速通道
            const hasQuestionItems = document.querySelectorAll('[class*="quest"], [class*="timu"], .paper-cnt, .paper-body').length > 0;
            const docCanvases = findDocumentPageCanvases(document.body);
            if (!hasQuestionItems && docCanvases.length >= 2) {
                const lca = getLowestCommonAncestor(docCanvases);
                if (lca && lca !== document.body && !isOwnPluginElement(lca)) {
                    return lca;
                }
            }

            let bestElement = null;
            let bestScore = 0;

            // 1. 先检查常见语义化正文、题库试卷与阅读器选择器
            for (const selector of CANDIDATE_SELECTORS) {
                const elements = document.querySelectorAll(selector);
                elements.forEach(el => {
                    if (isOwnPluginElement(el)) return;
                    const s = scoreElement(el) + 60;
                    if (s > bestScore) {
                        bestScore = s;
                        bestElement = el;
                    }
                });
            }

            // 2. 遍历页面所有 article / main / section / div / ul / ol 节点综合评分
            const allContainers = document.querySelectorAll('article, main, section, div, ul, ol');
            allContainers.forEach(el => {
                if (isOwnPluginElement(el)) return;
                const s = scoreElement(el);
                if (s > bestScore) {
                    bestScore = s;
                    bestElement = el;
                }
            });

            if (!bestElement) return document.body;

            // 3. 向内收敛 (Descendant Narrowing)：
            // 若当前选中的容器过宽（同时包住了左侧广告、中间试卷、右侧试题统计栏），自动向内收敛到承载核心试卷的子容器
            let narrowed = true;
            while (narrowed && bestElement) {
                narrowed = false;
                const children = Array.from(bestElement.children).filter(
                    c => !isOwnPluginElement(c) && isElementVisible(c) && !isEdgeSidebarRegion(c) && !isModalOrAnalysisDialog(c)
                );
                for (const child of children) {
                    const childScore = scoreElement(child);
                    if (childScore >= bestScore * 0.72 && childScore > 80) {
                        bestElement = child;
                        bestScore = childScore;
                        narrowed = true;
                        break;
                    }
                }
            }

            // 4. 向外扩展 (Ancestor Climb)：
            // 若当前只选中了某一道题或某一个大题分节，而其父容器包含全部大题且不是全屏多栏布局，则向上扩展一级
            let parent = bestElement.parentElement;
            while (parent && parent !== document.body && !isOwnPluginElement(parent)) {
                const parentRect = parent.getBoundingClientRect();
                if (window.innerWidth >= 1000 && parentRect.width > window.innerWidth * 0.88) {
                    break;
                }
                const parentScore = scoreElement(parent);
                const bestImgs = bestElement.querySelectorAll('img').length;
                const parentImgs = parent.querySelectorAll('img').length;
                const bestTextLen = (bestElement.innerText || '').trim().length;
                const parentTextLen = (parent.innerText || '').trim().length;
                if ((parentTextLen >= bestTextLen * 1.6 || parentImgs >= bestImgs * 1.6) && parentScore >= bestScore * 0.82) {
                    bestElement = parent;
                    bestScore = parentScore;
                    parent = bestElement.parentElement;
                } else {
                    break;
                }
            }

            return bestElement || document.body;
        } finally {
            if (wasOverlayShown) {
                overlayEl.style.display = '';
            }
        }
    }

    /**
     * 针对以整页 Canvas 画布为主的在线文档阅读器（如百度文库、PDF.js），提取所有已渲染（含滚动缓存）的纯净页面画布
     */
    function extractCanvasPagesAsContainer(docCanvases) {
        scanAndCacheVisibleCanvases();

        const cleanRoot = document.createElement('div');
        cleanRoot.className = `${APP_PREFIX}-extracted-content ${APP_PREFIX}-canvas-doc-mode`;

        // 收集所有大尺寸文档页 Canvas（宽 >= 380 且高 >= 380）的图像并按页面高度排序
        const pageItems = [];
        canvasBitmapCache.forEach((info, canvasEl) => {
            if (info.width >= 380 && info.height >= 380) {
                pageItems.push(info);
            }
        });

        let pageDataUrls = [];
        if (pageItems.length > 0) {
            pageDataUrls = pageItems.sort((a, b) => a.top - b.top).map(item => item.dataUrl);
        } else {
            pageDataUrls = docCanvases
                .map(c => getDrawnCanvasDataUrl(c) || (canvasBitmapCache.get(c) && canvasBitmapCache.get(c).dataUrl))
                .filter(Boolean);
        }

        pageDataUrls.forEach((dataUrl, idx) => {
            const pageBox = document.createElement('div');
            pageBox.className = `${APP_PREFIX}-canvas-page-item`;
            pageBox.style.margin = '0 auto 18px auto';
            pageBox.style.pageBreakInside = 'avoid';
            pageBox.style.breakInside = 'avoid';

            const img = document.createElement('img');
            img.src = dataUrl;
            img.alt = `文档第 ${idx + 1} 页`;
            img.className = `${APP_PREFIX}-block-img`;
            img.style.display = 'block';
            img.style.width = '100%';
            img.style.maxWidth = '100%';
            img.style.height = 'auto';
            img.style.border = '1px solid #e2e8f0';
            img.style.borderRadius = '4px';

            pageBox.appendChild(img);
            cleanRoot.appendChild(pageBox);
        });

        return cleanRoot;
    }

    /**
     * 提取并清洗目标 DOM 节点，保留排版、行内公式图片、独立大图、内嵌 iframe 文档框、表格、Canvas 与公式
     */
    function extractAndCleanNode(sourceNode, forceCanvasOnly = false, isManualSelection = false, iframeHtmlMap = null) {
        scanAndCacheVisibleCanvases();

        // 若目标节点本身就是一个文档预览 iframe，直接返回异步提取到的 iframe 内部文档内容
        const sourceTag = sourceNode.tagName ? sourceNode.tagName.toUpperCase() : '';
        if (sourceTag === 'IFRAME') {
            const cleanRoot = document.createElement('div');
            cleanRoot.className = `${APP_PREFIX}-extracted-content ${APP_PREFIX}-iframe-doc-mode`;
            const frameHtml = iframeHtmlMap && iframeHtmlMap.get(sourceNode);
            if (frameHtml) {
                cleanRoot.innerHTML = frameHtml;
            }
            return cleanRoot;
        }

        // 仅当用户手动点击“仅提画布页”，或目标容器是纯大画布阅读器（不含普通题库题目块且非用户手动框选）时，才走纯画布页模式
        const docCanvases = findDocumentPageCanvases(sourceNode);
        const hasHtmlQuestionBlocks = sourceNode.querySelectorAll &&
            sourceNode.querySelectorAll('[class*="quest"], [class*="timu"], [class*="exam-item"], table').length > 0;
        const isPureCanvasReader = forceCanvasOnly ||
            (!isManualSelection && !hasHtmlQuestionBlocks && docCanvases.length > 0 && (
                docCanvases.length >= 2 ||
                Boolean(sourceNode.closest && sourceNode.closest('[class*="creader"], #reader-container'))
            ));

        if (isPureCanvasReader) {
            const allPageCanvases = docCanvases.length > 0 ? docCanvases : findDocumentPageCanvases(document.body);
            if (allPageCanvases.length > 0) {
                return extractCanvasPagesAsContainer(allPageCanvases);
            }
        }

        // 若用户直接点选了单张 <img> 或单个 <canvas> 节点，直接包装导出，避免因无 firstChild 而变空白
        if (sourceTag === 'IMG') {
            const cleanRoot = document.createElement('div');
            cleanRoot.className = `${APP_PREFIX}-extracted-content`;
            const imgClone = sourceNode.cloneNode(true);
            if (imgClone.src) imgClone.setAttribute('src', imgClone.src);
            imgClone.classList.add(`${APP_PREFIX}-block-img`);
            cleanRoot.appendChild(imgClone);
            return cleanRoot;
        }
        if (sourceTag === 'CANVAS') {
            const cleanRoot = document.createElement('div');
            cleanRoot.className = `${APP_PREFIX}-extracted-content`;
            const dataUrl = getDrawnCanvasDataUrl(sourceNode) || (canvasBitmapCache.get(sourceNode) && canvasBitmapCache.get(sourceNode).dataUrl);
            if (dataUrl) {
                const img = document.createElement('img');
                img.src = dataUrl;
                img.className = `${APP_PREFIX}-block-img`;
                cleanRoot.appendChild(img);
            }
            return cleanRoot;
        }

        const clone = sourceNode.cloneNode(true);

        // 1. 最优先步骤：在尚未删除任何子节点前（此时 origCanvases 与 cloneCanvases 100% 一一对应），将所有已渲染或已缓存的 Canvas 转为 img 标签
        const origCanvases = sourceNode.querySelectorAll('canvas');
        const cloneCanvases = clone.querySelectorAll('canvas');
        if (origCanvases.length === cloneCanvases.length) {
            origCanvases.forEach((origCanvas, idx) => {
                const cloneCanvas = cloneCanvases[idx];
                const cached = canvasBitmapCache.get(origCanvas);
                const dataUrl = getDrawnCanvasDataUrl(origCanvas) || (cached && cached.dataUrl);
                if (dataUrl) {
                    const img = document.createElement('img');
                    img.src = dataUrl;
                    const rect = origCanvas.getBoundingClientRect();
                    if (rect.width >= 240 && rect.height >= 100) {
                        img.className = `${APP_PREFIX}-block-img`;
                        img.style.display = 'block';
                    } else {
                        img.style.display = 'inline-block';
                        img.style.verticalAlign = 'middle';
                    }
                    img.style.maxWidth = '100%';
                    img.style.height = 'auto';
                    let parent = cloneCanvas.parentElement;
                    while (parent && parent !== clone) {
                        if (parent.style && parent.style.height && parent.style.height.endsWith('px')) {
                            parent.style.height = 'auto';
                        }
                        parent = parent.parentElement;
                    }
                    cloneCanvas.replaceWith(img);
                } else {
                    cloneCanvas.remove();
                }
            });
        }

        // 1.1 在删除噪音节点前，优先将已提取到的文档预览 iframe 替换为真实 DOM 内容块，并解除其父容器固定高度滚动裁剪
        if (iframeHtmlMap && iframeHtmlMap.size > 0) {
            const origIframes = sourceNode.querySelectorAll('iframe');
            const cloneIframes = clone.querySelectorAll('iframe');
            if (origIframes.length === cloneIframes.length) {
                origIframes.forEach((origFrame, idx) => {
                    const frameHtml = iframeHtmlMap.get(origFrame);
                    if (frameHtml) {
                        const frameDiv = document.createElement('div');
                        frameDiv.className = `${APP_PREFIX}-iframe-content`;
                        frameDiv.innerHTML = frameHtml;
                        let parent = cloneIframes[idx].parentElement;
                        while (parent && parent !== clone) {
                            if (parent.style) {
                                parent.style.height = 'auto';
                                parent.style.maxHeight = 'none';
                                parent.style.overflow = 'visible';
                            }
                            parent = parent.parentElement;
                        }
                        cloneIframes[idx].replaceWith(frameDiv);
                    }
                });
            }
        }

        // 2. 标记原页面中的宽幅试题图片（宽 >= 200 且高 >= 22，例如组卷网的整道题图片条或几何配图），而将窄小的行内公式图片保持为 inline-block 垂直居中
        const origImgs = sourceNode.querySelectorAll('img');
        const cloneImgs = clone.querySelectorAll('img');
        if (origImgs.length === cloneImgs.length) {
            origImgs.forEach((origImg, idx) => {
                const rect = origImg.getBoundingClientRect();
                const cImg = cloneImgs[idx];
                if (rect.width >= 200 && rect.height >= 22) {
                    cImg.classList.add(`${APP_PREFIX}-block-img`);
                }
            });
        }

        // 2.1 识别无边框排版表格（如组卷网用于横向排列 A. B. C. D. 选项的 table），导出时去掉多余格线边框，同时保留真实数据表格的边框
        const origTables = sourceNode.querySelectorAll('table');
        const cloneTables = clone.querySelectorAll('table');
        if (origTables.length === cloneTables.length) {
            origTables.forEach((origTable, idx) => {
                const cTable = cloneTables[idx];
                const hasTh = Boolean(origTable.querySelector('th'));
                const firstTd = origTable.querySelector('td');
                const tdBorder = firstTd ? window.getComputedStyle(firstTd).borderTopWidth : '';
                const tableText = (origTable.innerText || '').trim();
                const isChoiceOptionTable = !hasTh && /[AＡ]\s*[.．、]/.test(tableText) && /[BＢ]\s*[.．、]/.test(tableText);
                const isBorderlessOriginal = !hasTh && (!tdBorder || tdBorder === '0px');
                if (isChoiceOptionTable || isBorderlessOriginal) {
                    cTable.classList.add(`${APP_PREFIX}-layout-table`);
                    cTable.style.setProperty('border', 'none', 'important');
                    cTable.querySelectorAll('tr, td, th').forEach(cell => {
                        cell.style.setProperty('border', 'none', 'important');
                        cell.style.setProperty('background', 'transparent', 'important');
                    });
                }
            });
        }

        // 3. 递归同步遍历源节点与克隆节点：仅移除显式隐藏的节点（display:none / visibility:hidden / opacity:0）、统计弹窗及整页容器中的独立侧边栏
        // 注意：绝不因 getBoundingClientRect() 尺寸小或处于左侧而误删正文内部的 <br>、公式 <img>、选项 <td> 等有效元素！
        const sourceRootRect = sourceNode.getBoundingClientRect();
        const isWholePageWideRoot = window.innerWidth >= 1000 && sourceRootRect.width >= window.innerWidth * 0.85;
        function syncCleanTree(srcEl, clnEl) {
            const srcChildren = Array.from(srcEl.children);
            const clnChildren = Array.from(clnEl.children);
            if (srcChildren.length !== clnChildren.length) return;

            for (let i = srcChildren.length - 1; i >= 0; i--) {
                const sChild = srcChildren[i];
                const cChild = clnChildren[i];

                if (isOwnPluginElement(sChild) || isModalOrAnalysisDialog(sChild)) {
                    cChild.remove();
                    continue;
                }

                const sTag = sChild.tagName.toUpperCase();
                // 始终保留换行标签 <br> 与图片/公式基础节点，避免因懒加载未完成或零宽而误删
                if (!['BR', 'IMG', 'SVG', 'MATH'].includes(sTag)) {
                    const style = window.getComputedStyle(sChild);
                    const isMathInternal = sChild.closest('.katex, .MathJax, mjx-container, math');
                    if (!isMathInternal && (style.display === 'none' || style.visibility === 'hidden' || parseFloat(style.opacity) === 0)) {
                        cChild.remove();
                        continue;
                    }
                }

                // 仅当源容器是占满全屏宽度的外层布局大容器时，才剔除其直接包含的大块左右侧边栏
                if (isWholePageWideRoot && isEdgeSidebarRegion(sChild)) {
                    cChild.remove();
                    continue;
                }

                // 若某个可见 div 采用 CSS background-image 渲染试题图片，自动补充为 <img> 标签以防打印丢失背景图
                if (sTag !== 'IMG' && sChild.children.length === 0) {
                    const bg = window.getComputedStyle(sChild).backgroundImage;
                    const sRect = sChild.getBoundingClientRect();
                    if (bg && bg.startsWith('url(') && sRect.width >= 150 && sRect.height >= 20) {
                        const urlMatch = bg.match(/url\(["']?([^"')]+)["']?\)/);
                        if (urlMatch && urlMatch[1]) {
                            const bgImg = document.createElement('img');
                            bgImg.src = urlMatch[1];
                            bgImg.className = `${APP_PREFIX}-block-img`;
                            cChild.appendChild(bgImg);
                        }
                    }
                }

                syncCleanTree(sChild, cChild);
            }
        }
        syncCleanTree(sourceNode, clone);

        // 4. 移除常见噪音元素（广告、评论、AI助手、搜索框、底部工具栏等）
        NOISE_SELECTORS.forEach(sel => {
            try {
                clone.querySelectorAll(sel).forEach(n => n.remove());
            } catch (_) {
                // 忽略无效选择器
            }
        });

        // 4.1 移除底部“剩余XX%未预览/未读、开通VIP、加入备课/挑错建议/版权异议、添加题型下试题、相关推荐列表”等非正文块
        const PROMO_BANNER_REGEX = /(剩余\s*\d+\s*[%页]\s*(未预览|未读)|开通\s*VIP\s*会员|新客免费下载|加入备课[\s\S]{0,40}(挑错建议|版权异议)|^添加题型下试题$|^显示全部答案和解析$)/i;
        const RECOMMEND_LIST_REGEX = /^(相关推荐|相关文档|猜你喜欢|热门推荐)/;
        Array.from(clone.querySelectorAll('div, section, aside, p, a, span')).forEach(el => {
            const txt = (el.innerText || el.textContent || '').trim();
            if (txt.length > 0 && txt.length < 220 && PROMO_BANNER_REGEX.test(txt) && !el.querySelector('table, article, img')) {
                el.remove();
                return;
            }
            if (txt.length > 0 && txt.length < 1500 && RECOMMEND_LIST_REGEX.test(txt) && /(阅读|下载)/.test(txt) && !el.querySelector(`.${APP_PREFIX}-iframe-content`)) {
                el.remove();
            }
        });

        // 5. 处理图片节点：修复懒加载属性 & 转换为绝对路径
        clone.querySelectorAll('img').forEach(img => {
            const lazySrc = img.getAttribute('data-src') ||
                            img.getAttribute('data-original') ||
                            img.getAttribute('data-actualsrc') ||
                            img.getAttribute('data-lazy-src');
            if ((!img.getAttribute('src') || img.getAttribute('src').startsWith('data:image/svg')) && lazySrc) {
                img.setAttribute('src', lazySrc);
            }
            if (img.src) {
                img.setAttribute('src', img.src);
            }
            img.removeAttribute('srcset');
            img.removeAttribute('loading');
            img.style.maxWidth = '100%';
            img.style.height = 'auto';
        });

        // 6. 清理干扰排版的危险内联样式及题目外层卡片框线/阴影（如固定定位、负边距、卡片边框、固定高度截断等）
        clone.querySelectorAll('*').forEach(el => {
            el.removeAttribute('onclick');
            el.removeAttribute('onmousedown');
            el.removeAttribute('onselectstart');
            if (el.style) {
                if (el.style.position === 'fixed' || el.style.position === 'sticky') {
                    el.style.position = 'static';
                }
                el.style.maxHeight = 'none';
                el.style.overflow = 'visible';
                el.style.userSelect = 'auto';
                if (['DIV', 'SECTION', 'ARTICLE', 'LI', 'UL'].includes(el.tagName)) {
                    if (el.style.height && el.style.height.endsWith('px')) {
                        el.style.height = 'auto';
                    }
                    el.style.setProperty('border', 'none', 'important');
                    el.style.setProperty('box-shadow', 'none', 'important');
                    el.style.setProperty('background', 'transparent', 'important');
                    el.style.setProperty('border-radius', '0', 'important');
                }
            }
        });

        // 7. 重置克隆根节点本身的内外边距与宽高约束，避免与纸张容器双重叠加或导致分页截断
        const cleanRoot = document.createElement('div');
        cleanRoot.className = `${APP_PREFIX}-extracted-content`;
        if (clone.childNodes.length > 0) {
            while (clone.firstChild) {
                cleanRoot.appendChild(clone.firstChild);
            }
        } else {
            cleanRoot.appendChild(clone);
        }

        return cleanRoot;
    }

    /**
     * 读取元素的 ::before 或 ::after 伪元素文本内容
     */
    function getPseudoElementText(el, pseudoSelector) {
        if (!el || el.nodeType !== Node.ELEMENT_NODE) return '';
        try {
            const style = window.getComputedStyle(el, pseudoSelector);
            if (!style) return '';
            const rawContent = style.getPropertyValue('content');
            if (!rawContent || rawContent === 'none' || rawContent === 'normal') return '';
            // 去除首尾引号
            const cleaned = rawContent.replace(/^["']+|["']+$/g, '').trim();
            if (!cleaned) return '';
            // 过滤掉 counter() / url() 函数表达式以及纯图标字体私有区字符 (Private Use Area)
            if (/^(url|counter|counters|attr)\s*\(/i.test(rawContent)) return '';
            if (/^[\uE000-\uF8FF\s]+$/.test(cleaned)) return '';
            return cleaned;
        } catch (_) {
            return '';
        }
    }

    /**
     * HTML 实体转义辅助函数
     */
    function escapeHtmlText(str) {
        return String(str)
            .replace(/&/g, '&amp;')
            .replace(/</g, '&lt;')
            .replace(/>/g, '&gt;');
    }

    /**
     * 深度全量 DOM 提取函数：
     * 遍历目标根节点（默认整个 document.body 或正文容器），找出所有包含文字、伪元素 (::before / ::after) 或图片的节点，
     * 无论它们当前是否被 CSS 隐藏（如 display: none、visibility: hidden、opacity: 0），
     * 将其文本内容、伪元素内容和图片地址按原 DOM 顺序拼接成一段完整的 HTML 字符串返回。
     *
     * @param {Element} rootNode - 要遍历的根节点，默认为 document.body
     * @param {boolean} filterNoise - 是否跳过导航栏/广告/侧边栏等明显噪音容器，默认 true
     * @returns {string} 按原顺序拼接好的完整 HTML 字符串
     */
    function extractAllDomWithHiddenAndPseudo(rootNode = document.body, filterNoise = true, iframeHtmlMap = null) {
        scanAndCacheVisibleCanvases();

        if (rootNode && rootNode.tagName === 'IFRAME' && iframeHtmlMap && iframeHtmlMap.get(rootNode)) {
            return iframeHtmlMap.get(rootNode);
        }

        const SKIP_TAGS = new Set([
            'SCRIPT', 'STYLE', 'NOSCRIPT', 'TEMPLATE',
            'OBJECT', 'EMBED', 'HEAD', 'META', 'LINK', 'BUTTON',
            'INPUT', 'SELECT', 'TEXTAREA'
        ]);

        const BLOCK_TAGS = new Set([
            'ARTICLE', 'MAIN', 'SECTION', 'DIV', 'P', 'H1', 'H2', 'H3', 'H4', 'H5', 'H6',
            'UL', 'OL', 'LI', 'DL', 'DT', 'DD', 'BLOCKQUOTE', 'PRE', 'HEADER', 'FIGCAPTION', 'FIGURE'
        ]);

        const htmlChunks = [];

        function serializeNode(node) {
            if (!node) return '';

            // 1. 文本节点：无论父节点是否 display:none，均提取其文本内容
            if (node.nodeType === Node.TEXT_NODE) {
                const txt = node.nodeValue ? node.nodeValue.replace(/\s+/g, ' ') : '';
                return txt.trim() ? escapeHtmlText(txt) : '';
            }

            if (node.nodeType !== Node.ELEMENT_NODE) return '';

            const el = node;
            const tag = el.tagName.toUpperCase();

            // 跳过脚本、样式及本插件自身注入的 UI 节点
            if (SKIP_TAGS.has(tag) || isOwnPluginElement(el)) return '';

            // 1.1 若为文档预览 iframe，直接内联其提取结果
            if (tag === 'IFRAME') {
                const frameHtml = iframeHtmlMap && iframeHtmlMap.get(el);
                return frameHtml ? `<div class="${APP_PREFIX}-iframe-content">${frameHtml}</div>` : '';
            }

            // 可选跳过明显的导航、广告、评论、侧边栏噪音块及“试卷分析/细目表”统计弹窗
            if (filterNoise) {
                if (['NAV', 'ASIDE', 'FOOTER'].includes(tag)) return '';
                const classAndId = `${typeof el.className === 'string' ? el.className : ''} ${el.id || ''}`.toLowerCase();
                if (NOISE_CLASS_REGEX.test(classAndId)) return '';
                if (isModalOrAnalysisDialog(el)) return '';
                if (el !== rootNode && isEdgeSidebarRegion(el)) return '';
            }

            // 2. 处理图片节点 (<img>)：支持懒加载属性，区分行内公式小图与宽幅试题图/独立配图
            if (tag === 'IMG') {
                const rawSrc = el.getAttribute('data-src') ||
                               el.getAttribute('data-original') ||
                               el.getAttribute('data-actualsrc') ||
                               el.getAttribute('data-lazy-src') ||
                               el.getAttribute('src') ||
                               el.src;
                if (!rawSrc) return '';
                const tempImg = document.createElement('img');
                tempImg.src = rawSrc;
                const absSrc = tempImg.src || rawSrc;
                const alt = escapeHtmlText(el.getAttribute('alt') || '');
                const rect = el.getBoundingClientRect();
                const isLargeBlock = rect.width >= 200 && rect.height >= 22;
                const styleStr = isLargeBlock
                    ? 'max-width:100%;height:auto;display:block;margin:10px 0;'
                    : 'max-width:100%;height:auto;display:inline-block;vertical-align:middle;margin:0 2px;';
                const clsStr = isLargeBlock ? ` class="${APP_PREFIX}-block-img"` : '';
                return `<img src="${absSrc}" alt="${alt}"${clsStr} style="${styleStr}" />`;
            }

            // 3. 处理画布节点 (<canvas>)：优先读取当前画面或滚动缓存中的图像
            if (tag === 'CANVAS') {
                const cached = canvasBitmapCache.get(el);
                const dataUrl = getDrawnCanvasDataUrl(el) || (cached && cached.dataUrl);
                if (dataUrl) {
                    return `<img src="${dataUrl}" alt="canvas-image" class="${APP_PREFIX}-block-img" style="max-width:100%;height:auto;display:block;margin:12px auto;" />`;
                }
                return '';
            }

            // 4. 处理数学公式节点 (KaTeX / MathJax / MathML) 或内联 SVG 图形
            if (tag === 'MATH' || tag === 'MJX-CONTAINER' || el.classList.contains('katex') || el.classList.contains('MathJax')) {
                return el.outerHTML;
            }
            if (tag === 'SVG') {
                const w = el.getAttribute('width') || el.getBoundingClientRect().width;
                const h = el.getAttribute('height') || el.getBoundingClientRect().height;
                if (w > 10 && h > 10) {
                    return el.outerHTML;
                }
                return '';
            }

            // 5. 处理表格 (<table>)：完整保留表格结构，识别无边框选择题选项排版表并去掉框线
            if (tag === 'TABLE') {
                const tableClone = el.cloneNode(true);
                const hasTh = Boolean(el.querySelector('th'));
                const firstTd = el.querySelector('td');
                const tdBorder = firstTd ? window.getComputedStyle(firstTd).borderTopWidth : '';
                const tableText = (el.innerText || '').trim();
                const isLayoutTable = (!hasTh && /[AＡ]\s*[.．、]/.test(tableText) && /[BＢ]\s*[.．、]/.test(tableText)) || (!hasTh && (!tdBorder || tdBorder === '0px'));
                if (isLayoutTable) {
                    tableClone.classList.add(`${APP_PREFIX}-layout-table`);
                    tableClone.style.setProperty('border', 'none', 'important');
                }
                tableClone.querySelectorAll('*').forEach(cell => {
                    if (cell.style) {
                        if (cell.style.display === 'none') cell.style.display = '';
                        if (cell.style.visibility === 'hidden') cell.style.visibility = 'visible';
                        if (cell.style.opacity === '0') cell.style.opacity = '1';
                        if (isLayoutTable && ['TR', 'TD', 'TH'].includes(cell.tagName)) {
                            cell.style.setProperty('border', 'none', 'important');
                            cell.style.setProperty('background', 'transparent', 'important');
                        }
                    }
                });
                return tableClone.outerHTML;
            }

            // 6. 提取 ::before 与 ::after 伪元素插入的文本内容
            const beforeText = getPseudoElementText(el, '::before');
            const afterText = getPseudoElementText(el, '::after');

            // 7. 递归遍历所有子节点（按原顺序拼接）
            let childrenHtml = '';
            const childNodes = el.childNodes;
            for (let i = 0; i < childNodes.length; i++) {
                const part = serializeNode(childNodes[i]);
                if (part) {
                    childrenHtml += (childrenHtml && !childrenHtml.endsWith('>') && !part.startsWith('<') ? ' ' : '') + part;
                }
            }

            const combinedInner = `${beforeText ? `<span class="${APP_PREFIX}-pseudo-before">${escapeHtmlText(beforeText)}</span>` : ''}${childrenHtml}${afterText ? `<span class="${APP_PREFIX}-pseudo-after">${escapeHtmlText(afterText)}</span>` : ''}`.trim();

            if (!combinedInner) return '';

            // 8. 按原语义标签包裹输出，保证段落、标题与列表层级不乱
            if (/^H[1-6]$/.test(tag)) {
                const lower = tag.toLowerCase();
                return `<${lower}>${combinedInner}</${lower}>`;
            }
            if (tag === 'BLOCKQUOTE' || tag === 'PRE') {
                const lower = tag.toLowerCase();
                return `<${lower}>${combinedInner}</${lower}>`;
            }
            if (tag === 'UL' || tag === 'OL') {
                const lower = tag.toLowerCase();
                return `<${lower}>${combinedInner}</${lower}>`;
            }
            if (tag === 'LI') {
                return `<li>${combinedInner}</li>`;
            }
            if (tag === 'BR') {
                return '<br/>';
            }
            if (tag === 'STRONG' || tag === 'B') {
                return `<strong>${combinedInner}</strong>`;
            }
            if (tag === 'EM' || tag === 'I') {
                return `<em>${combinedInner}</em>`;
            }
            if (tag === 'SUB' || tag === 'SUP' || tag === 'CODE') {
                const lower = tag.toLowerCase();
                return `<${lower}>${combinedInner}</${lower}>`;
            }
            if (BLOCK_TAGS.has(tag)) {
                if (/^<(div|p|h[1-6]|ul|ol|li|table|blockquote|pre)\b/i.test(combinedInner)) {
                    return `<div class="${APP_PREFIX}-deep-block">${combinedInner}</div>`;
                }
                return `<p>${combinedInner}</p>`;
            }

            return `<span>${combinedInner}</span>`;
        }

        const bodyHtml = serializeNode(rootNode);
        if (bodyHtml) {
            htmlChunks.push(bodyHtml);
        }

        return htmlChunks.join('\n');
    }

    // 将深度提取函数挂载到 window，方便用户在控制台或其他脚本中直接调用
    window.extractAllDomWithHiddenAndPseudo = extractAllDomWithHiddenAndPseudo;

    // ==================== 3. 跨域图片转 Base64 (CORS 兼容) ====================

    /**
     * 将单个图片 URL 转为 Base64 DataURL
     * 1. 先尝试页面上下文同源/带 Cookie 的 fetch（支持 blob: URL 与 alicfw 防火墙 Cookie）
     * 2. 再使用油猴 GM_xmlhttpRequest（携带 Referer 与 Cookie，并校验返回 MIME 必须为 image/*，防止将防火墙 HTML 页面转成损坏的 data:text/html）
     * 3. 最后回退至 Canvas 转换
     */
    function imageUrlToBase64(url) {
        return new Promise((resolve) => {
            if (!url) {
                resolve(url);
                return;
            }
            if (url.startsWith('data:image/png') || url.startsWith('data:image/jpeg') || url.startsWith('data:image/jpg') || url.startsWith('data:image/webp') || url.startsWith('data:image/gif')) {
                resolve(url);
                return;
            }
            if (url.startsWith('data:image/svg')) {
                fallbackCanvasConvert(url).then(resolve);
                return;
            }

            const blobToValidImageDataUrl = (blob, onFail) => {
                if (!blob || (blob.type && !blob.type.startsWith('image/'))) {
                    onFail();
                    return;
                }
                const reader = new FileReader();
                reader.onloadend = () => {
                    const res = reader.result;
                    if (typeof res === 'string' && res.startsWith('data:image/')) {
                        resolve(res);
                    } else {
                        onFail();
                    }
                };
                reader.onerror = onFail;
                reader.readAsDataURL(blob);
            };

            const tryGmRequest = () => {
                if (typeof GM_xmlhttpRequest === 'function' && !url.startsWith('blob:')) {
                    GM_xmlhttpRequest({
                        method: 'GET',
                        url: url,
                        headers: {
                            'Referer': window.location.href,
                            'Origin': window.location.origin
                        },
                        anonymous: false,
                        responseType: 'blob',
                        timeout: 8000,
                        onload: function (response) {
                            if (response.status >= 200 && response.status < 300 && response.response) {
                                blobToValidImageDataUrl(response.response, () => fallbackCanvasConvert(url).then(resolve));
                            } else {
                                fallbackCanvasConvert(url).then(resolve);
                            }
                        },
                        onerror: function () {
                            fallbackCanvasConvert(url).then(resolve);
                        },
                        ontimeout: function () {
                            fallbackCanvasConvert(url).then(resolve);
                        }
                    });
                } else {
                    fallbackCanvasConvert(url).then(resolve);
                }
            };

            // 优先在当前页面上下文带凭证 fetch（兼容 blob: URL 及带防火墙 Cookie 的同站图片）
            fetch(url, { credentials: 'include' })
                .then(res => {
                    if (!res.ok) throw new Error('fetch failed');
                    return res.blob();
                })
                .then(blob => {
                    blobToValidImageDataUrl(blob, tryGmRequest);
                })
                .catch(() => {
                    tryGmRequest();
                });
        });
    }

    /**
     * 方案 B: 使用 crossOrigin="anonymous" + Canvas 转换
     */
    function fallbackCanvasConvert(url) {
        return new Promise((resolve) => {
            const img = new Image();
            if (!url.startsWith('data:')) {
                img.crossOrigin = 'anonymous';
            }
            const timer = setTimeout(() => resolve(url), 6000);

            img.onload = function () {
                clearTimeout(timer);
                try {
                    const canvas = document.createElement('canvas');
                    canvas.width = img.naturalWidth || img.width || 520;
                    canvas.height = img.naturalHeight || img.height || 240;
                    const ctx = canvas.getContext('2d');
                    ctx.fillStyle = '#ffffff';
                    ctx.fillRect(0, 0, canvas.width, canvas.height);
                    ctx.drawImage(img, 0, 0);
                    resolve(canvas.toDataURL('image/png'));
                } catch (_) {
                    resolve(url);
                }
            };
            img.onerror = function () {
                clearTimeout(timer);
                resolve(url);
            };
            img.src = url;
        });
    }

    /**
     * 批量将容器内的所有图片内联为 Base64
     */
    async function inlineAllImagesAsBase64(container) {
        const imgs = Array.from(container.querySelectorAll('img'));
        if (imgs.length > 0) {
            await Promise.all(imgs.map(async (img) => {
                const src = img.getAttribute('src') || img.src;
                if (!src) return;
                const base64 = await imageUrlToBase64(src);
                img.setAttribute('src', base64);
            }));
        }

        // 同时将 SVG 转为兼容显示或保持内联尺寸
        container.querySelectorAll('svg').forEach(svg => {
            const rect = svg.getBoundingClientRect();
            if (rect.width > 0 && !svg.getAttribute('width')) {
                svg.setAttribute('width', Math.round(rect.width));
            }
            if (rect.height > 0 && !svg.getAttribute('height')) {
                svg.setAttribute('height', Math.round(rect.height));
            }
        });
    }

    // ==================== 4. 多格式导出功能 (PDF / Word / 打印) ====================

    function getSafeFilename(ext) {
        const rawTitle = (document.title || '网页提取资料').trim();
        const safeTitle = rawTitle.replace(/[\\/:*?"<>|]/g, '_').slice(0, 60);
        return `${safeTitle}.${ext}`;
    }

    /**
     * 4.1 导出为 PDF (基于 html2pdf.js)
     * 关键修复：
     * 1. 创建独立于 fixed 弹窗之外的导出节点（.cre-exporter-pdf-render-box），避免被 position:fixed 和 body overflow:hidden 裁剪为单屏高度。
     * 2. 传入 scrollX: 0, scrollY: 0，防止页面滚动条偏移导致顶部留白、底部内容被截断。
     * 3. 移除 pagebreak 中的 'avoid-all'（避免整个正文容器被当成不可分割块导致跨页截断），仅对表格行、图片、引用块启用防截断。
     */
    async function exportToPDF(paperEl) {
        showToast('⏳ 正在处理图片与多页排版，准备生成 PDF...', 5000);
        await inlineAllImagesAsBase64(paperEl);

        if (typeof html2pdf === 'undefined') {
            showToast('⚠️ html2pdf 库加载受阻，已为您切换至系统高清打印存为 PDF');
            triggerSystemPrint(paperEl);
            return;
        }

        // 构造专门用于 PDF 渲染的纯净节点（避免与 #cre-exporter-paper ID 冲突及双重内边距）
        const pdfRenderNode = document.createElement('div');
        pdfRenderNode.className = `${APP_PREFIX}-pdf-render-box`;
        pdfRenderNode.innerHTML = paperEl.innerHTML;

        // 临时解除 body 的 overflow:hidden 并隐藏 fixed 遮罩层，防止 html2canvas 克隆文档时将画布裁剪在 100vh 首屏高度
        const savedScrollX = window.scrollX;
        const savedScrollY = window.scrollY;
        const wasOverlayActive = overlayEl && overlayEl.classList.contains('active');

        if (wasOverlayActive) {
            overlayEl.style.display = 'none';
        }
        document.body.style.overflow = 'visible';
        window.scrollTo(0, 0);

        const opt = {
            margin: [12, 12, 12, 12],
            filename: getSafeFilename('pdf'),
            image: { type: 'jpeg', quality: 0.96 },
            html2canvas: {
                scale: 2,
                useCORS: true,
                allowTaint: true,
                logging: false,
                backgroundColor: '#ffffff',
                scrollX: 0,
                scrollY: 0
            },
            jsPDF: {
                unit: 'mm',
                format: 'a4',
                orientation: 'portrait'
            },
            pagebreak: {
                mode: ['css', 'legacy'],
                avoid: ['p', 'li', 'tr', 'img', 'blockquote', 'pre', 'h1', 'h2', 'h3', 'h4', `.${APP_PREFIX}-canvas-page-item`]
            }
        };

        try {
            await html2pdf().set(opt).from(pdfRenderNode).save();
            showToast('✅ PDF 完整导出成功！');
        } catch (err) {
            console.error('[CleanReader] PDF 导出异常:', err);
            showToast('⚠️ 前端渲染 PDF 受限，正在调用系统“另存为 PDF”...');
            if (wasOverlayActive) {
                overlayEl.style.display = '';
                document.body.style.overflow = 'hidden';
            }
            triggerSystemPrint(paperEl);
            return;
        } finally {
            if (wasOverlayActive) {
                overlayEl.style.display = '';
                document.body.style.overflow = 'hidden';
            }
            window.scrollTo(savedScrollX, savedScrollY);
        }
    }

    /**
     * 4.2 导出为 Word 文档 (内嵌 Base64 图片与完整排版样式)
     */
    async function exportToWord(paperEl) {
        showToast('⏳ 正在转换图片并生成 Word 文档...');
        await inlineAllImagesAsBase64(paperEl);

        const clone = paperEl.cloneNode(true);

        // 限制图片在 Word 页面中的最大宽度（A4 页面内容宽度约 550pt）
        clone.querySelectorAll('img').forEach(img => {
            const w = img.naturalWidth || img.width;
            if (w > 560) {
                img.setAttribute('width', '560');
            } else if (w > 0) {
                img.setAttribute('width', String(w));
            }
            img.style.maxWidth = '100%';
            img.style.height = 'auto';
        });

        const wordDocumentHTML = `
            <html xmlns:o="urn:schemas-microsoft-com:office:office"
                  xmlns:w="urn:schemas-microsoft-com:office:word"
                  xmlns:m="http://schemas.microsoft.com/office/2004/12/omml"
                  xmlns="http://www.w3.org/TR/REC-html40">
            <head>
                <meta charset="utf-8">
                <title>${document.title || '导出文档'}</title>
                <!--[if gte mso 9]>
                <xml>
                    <w:WordDocument>
                        <w:View>Print</w:View>
                        <w:Zoom>100</w:Zoom>
                        <w:DoNotOptimizeForBrowser/>
                    </w:WordDocument>
                </xml>
                <![endif]-->
                <style>
                    @page {
                        size: 21cm 29.7cm;
                        margin: 2.2cm 2.0cm 2.2cm 2.0cm;
                    }
                    body {
                        font-family: "SimSun", "Songti SC", "Microsoft YaHei", Arial, sans-serif;
                        font-size: 11.5pt;
                        line-height: 1.65;
                        color: #111827;
                    }
                    h1 { font-size: 18pt; font-weight: bold; color: #0f172a; border-bottom: 1px solid #cbd5e1; padding-bottom: 6pt; margin-bottom: 12pt; }
                    h2 { font-size: 15pt; font-weight: bold; color: #0f172a; margin-top: 14pt; margin-bottom: 6pt; }
                    h3 { font-size: 13pt; font-weight: bold; color: #0f172a; margin-top: 10pt; margin-bottom: 4pt; }
                    p { margin: 0 0 8pt 0; }
                    table {
                        width: 100%;
                        border-collapse: collapse;
                        margin: 10pt 0;
                    }
                    th, td {
                        border: 1pt solid #94a3b8;
                        padding: 6pt 8pt;
                        font-size: 10.5pt;
                    }
                    table.${APP_PREFIX}-layout-table,
                    table.${APP_PREFIX}-layout-table tr,
                    table.${APP_PREFIX}-layout-table td,
                    table.${APP_PREFIX}-layout-table th {
                        border: none !important;
                        padding: 3pt 8pt 3pt 0 !important;
                        background: transparent !important;
                    }
                    th {
                        background-color: #f1f5f9;
                        font-weight: bold;
                    }
                    pre, blockquote {
                        background-color: #f8fafc;
                        border-left: 3pt solid #3b82f6;
                        padding: 8pt 10pt;
                        margin: 8pt 0;
                    }
                    img {
                        max-width: 100%;
                        height: auto;
                    }
                </style>
            </head>
            <body>
                ${clone.innerHTML}
            </body>
            </html>
        `;

        // 若页面加载了 htmlDocx 则生成标准 .docx，否则生成带 MSO XML 头的富格式 Word 文件
        let blob;
        let filename;
        if (window.htmlDocx && typeof window.htmlDocx.asBlob === 'function') {
            blob = window.htmlDocx.asBlob(wordDocumentHTML);
            filename = getSafeFilename('docx');
        } else {
            blob = new Blob(['\ufeff', wordDocumentHTML], {
                type: 'application/msword;charset=utf-8'
            });
            filename = getSafeFilename('doc');
        }

        const link = document.createElement('a');
        link.href = URL.createObjectURL(blob);
        link.download = filename;
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
        setTimeout(() => URL.revokeObjectURL(link.href), 5000);

        showToast(`✅ Word 文档 (${filename}) 已导出！`);
    }

    /**
     * 4.3 调用独立 iframe 隔离打印视图 / 另存为矢量高清 PDF
     * 说明：将提取后的纯净 DOM 克隆到一个全新的空白 iframe 中，
     * 注入独立的 A4 打印专用 CSS，彻底隔离原网页复杂/错乱的样式表，
     * 再调用 iframe.contentWindow.print() 触发系统打印或另存为 PDF。
     */
    async function triggerSystemPrint(sourcePaperEl) {
        const targetPaper = sourcePaperEl || paperEl;
        if (!targetPaper) return;

        showToast('⏳ 正在构建独立 A4 打印视图...');
        await inlineAllImagesAsBase64(targetPaper);

        // 移除旧的打印 iframe（若存在）
        const oldIframe = document.getElementById(`${APP_PREFIX}-print-iframe`);
        if (oldIframe) oldIframe.remove();

        const printIframe = document.createElement('iframe');
        printIframe.id = `${APP_PREFIX}-print-iframe`;
        printIframe.style.position = 'fixed';
        printIframe.style.left = '-9999px';
        printIframe.style.top = '0';
        printIframe.style.width = '800px';
        printIframe.style.height = '1120px';
        printIframe.style.border = '0';
        printIframe.style.opacity = '0';
        printIframe.style.pointerEvents = 'none';
        document.body.appendChild(printIframe);

        const iframeDoc = printIframe.contentWindow.document;
        iframeDoc.open();
        iframeDoc.write(`
            <!DOCTYPE html>
            <html lang="zh-CN">
            <head>
                <meta charset="UTF-8">
                <title>${document.title || '打印文档'}</title>
                <style>
                    /* A4 纸张打印排版与无障碍阅读重置样式 */
                    @page {
                        size: A4 portrait;
                        margin: 18mm 16mm 18mm 16mm;
                    }
                    * {
                        box-sizing: border-box;
                        -webkit-print-color-adjust: exact !important;
                        print-color-adjust: exact !important;
                    }
                    html, body {
                        margin: 0;
                        padding: 0;
                        background: #ffffff !important;
                        color: #111827 !important;
                        font-family: -apple-system, BlinkMacSystemFont, "PingFang SC", "Microsoft YaHei", "Noto Serif SC", serif;
                        font-size: 11.5pt;
                        line-height: 1.75;
                        user-select: text;
                    }
                    h1, h2, h3, h4 {
                        color: #0f172a;
                        font-weight: 700;
                        line-height: 1.35;
                        page-break-after: avoid;
                        break-after: avoid;
                    }
                    h1 {
                        font-size: 19pt;
                        margin: 0 0 12pt 0;
                        padding-bottom: 8pt;
                        border-bottom: 1.5pt solid #cbd5e1;
                    }
                    h2 { font-size: 15pt; margin: 16pt 0 8pt 0; }
                    h3 { font-size: 13pt; margin: 12pt 0 6pt 0; }
                    p {
                        margin: 0 0 9pt 0;
                        text-align: justify;
                        orphans: 3;
                        widows: 3;
                    }
                    img, svg, canvas {
                        max-width: 100% !important;
                        height: auto !important;
                        display: block;
                        margin: 10pt auto;
                        page-break-inside: avoid;
                        break-inside: avoid;
                    }
                    table {
                        width: 100%;
                        border-collapse: collapse;
                        margin: 12pt 0;
                        font-size: 10.5pt;
                        page-break-inside: avoid;
                        break-inside: avoid;
                    }
                    th, td {
                        border: 1pt solid #94a3b8;
                        padding: 6pt 10pt;
                        text-align: left;
                        vertical-align: top;
                    }
                    table.${APP_PREFIX}-layout-table,
                    table.${APP_PREFIX}-layout-table tr,
                    table.${APP_PREFIX}-layout-table td,
                    table.${APP_PREFIX}-layout-table th {
                        border: none !important;
                        padding: 3pt 8pt 3pt 0 !important;
                        background: transparent !important;
                    }
                    th {
                        background-color: #f1f5f9 !important;
                        font-weight: 600;
                    }
                    blockquote, pre {
                        background: #f8fafc !important;
                        border-left: 3pt solid #2563eb;
                        padding: 8pt 12pt;
                        margin: 10pt 0;
                        border-radius: 3pt;
                        page-break-inside: avoid;
                        break-inside: avoid;
                        white-space: pre-wrap;
                        word-wrap: break-word;
                    }
                </style>
            </head>
            <body>
                <div class="a4-print-container">
                    ${targetPaper.innerHTML}
                </div>
            </body>
            </html>
        `);
        iframeDoc.close();

        // 等待 iframe 内部图片与 DOM 渲染就绪后调起系统打印
        setTimeout(() => {
            try {
                printIframe.contentWindow.focus();
                printIframe.contentWindow.print();
            } catch (err) {
                console.error('[CleanReader] iframe 打印调用失败，回退至当前页打印:', err);
                document.body.classList.add(`${APP_PREFIX}-printing`);
                window.print();
                document.body.classList.remove(`${APP_PREFIX}-printing`);
            }
        }, 250);
    }

    // ==================== 5. 界面与交互构建 ====================

    let overlayEl = null;
    let paperEl = null;
    let toastEl = null;
    let toastTimer = null;
    let isPickerActive = false;
    let hoveredTarget = null;
    let currentSourceNode = null;

    function showToast(msg, duration = 3200) {
        if (!toastEl) return;
        toastEl.textContent = msg;
        toastEl.style.display = 'block';
        clearTimeout(toastTimer);
        toastTimer = setTimeout(() => {
            toastEl.style.display = 'none';
        }, duration);
    }

    /**
     * 打开纯净阅读视图并填充提取的内容（支持同源与跨源文档预览 iframe 自动异步提取）
     */
    async function openCleanView(customTargetNode = null, forceCanvasOnly = false) {
        const isManual = Boolean(customTargetNode);
        const targetNode = customTargetNode || detectMainContentNode();
        currentSourceNode = targetNode;

        // 异步提取目标节点内部（或目标本身）的文档预览 iframe 内容
        const iframeHtmlMap = await collectContentIframes(targetNode, false);
        const cleanedNode = extractAndCleanNode(targetNode, forceCanvasOnly, isManual, iframeHtmlMap);

        paperEl.innerHTML = '';

        // 若提取区域中没有主标题 h1，自动补上页面标题
        if (!cleanedNode.querySelector('h1') && document.title) {
            const titleEl = document.createElement('h1');
            titleEl.textContent = document.title;
            paperEl.appendChild(titleEl);
        }

        paperEl.appendChild(cleanedNode);
        overlayEl.classList.add('active');
        document.body.classList.add(`${APP_PREFIX}-clean-open`);
        document.body.style.overflow = 'hidden';

        const isCanvasDoc = cleanedNode.classList.contains(`${APP_PREFIX}-canvas-doc-mode`);
        const isIframeDoc = iframeHtmlMap.size > 0;
        const badgeEl = overlayEl.querySelector(`.${APP_PREFIX}-toolbar-badge`);
        if (badgeEl) {
            if (isIframeDoc) {
                badgeEl.textContent = '已自动穿透提取内嵌滚动文档框 (iframe) 全部内容';
            } else if (isCanvasDoc) {
                const count = cleanedNode.querySelectorAll('img').length;
                badgeEl.textContent = `已提取 ${count} 页文档画布（滚动原网页加载更多页后可重提）`;
            } else if (isManual) {
                badgeEl.textContent = '已提取您手动选定的区域（点击“🔼 扩大范围”可选上一层父容器）';
            } else {
                badgeEl.textContent = '已过滤侧边栏、广告与无关导航';
            }
        }

        return paperEl;
    }

    /**
     * 开启深度全量提取视图（包含被 CSS display:none / visibility:hidden 隐藏的节点、::before/::after 伪元素及内嵌 iframe 文档框）
     */
    async function openDeepExtractView(customRoot = null) {
        const wasOverlayShown = overlayEl && overlayEl.classList.contains('active') && overlayEl.style.display !== 'none';
        if (wasOverlayShown) {
            overlayEl.style.display = 'none';
        }

        let targetNode = customRoot;
        if (!targetNode) {
            const detected = detectMainContentNode();
            targetNode = detected && detected !== document.body ? detected : document.body;
        }
        currentSourceNode = targetNode;

        const iframeHtmlMap = await collectContentIframes(targetNode, true);
        const fullHtmlString = extractAllDomWithHiddenAndPseudo(targetNode, true, iframeHtmlMap);

        if (wasOverlayShown) {
            overlayEl.style.display = '';
        }

        paperEl.innerHTML = '';
        const wrapper = document.createElement('div');
        wrapper.className = `${APP_PREFIX}-extracted-content ${APP_PREFIX}-deep-extract-mode`;
        wrapper.innerHTML = fullHtmlString;

        if (!wrapper.querySelector('h1') && document.title) {
            const titleEl = document.createElement('h1');
            titleEl.textContent = document.title;
            paperEl.appendChild(titleEl);
        }

        paperEl.appendChild(wrapper);
        overlayEl.classList.add('active');
        document.body.classList.add(`${APP_PREFIX}-clean-open`);
        document.body.style.overflow = 'hidden';

        const badgeEl = overlayEl.querySelector(`.${APP_PREFIX}-toolbar-badge`);
        if (badgeEl) {
            badgeEl.textContent = '深度提取模式：已还原 CSS 隐藏节点、伪元素与内嵌 iframe 文档框';
        }

        return paperEl;
    }

    function closeCleanView() {
        overlayEl.classList.remove('active');
        document.body.classList.remove(`${APP_PREFIX}-clean-open`);
        paperEl.setAttribute('contenteditable', 'false');
        document.body.style.overflow = '';
    }

    /**
     * 启动手动点选网页区域模式（支持智能容器吸附 + 键盘上下键扩大/缩小选区范围）
     */
    function startAreaPicker() {
        if (isPickerActive) return;
        isPickerActive = true;
        showToast('🎯 请移动鼠标框选试卷/正文区域，左键点击确认（按 ↑ 键扩大到父容器，ESC 取消）', 7000);

        function highlightElement(el) {
            if (!el || isOwnPluginElement(el)) return;
            if (hoveredTarget && hoveredTarget !== el) {
                hoveredTarget.classList.remove(`${APP_PREFIX}-picker-highlight`);
            }
            hoveredTarget = el;
            hoveredTarget.classList.add(`${APP_PREFIX}-picker-highlight`);
        }

        function onMouseMove(e) {
            let el = e.target;
            if (!el || isOwnPluginElement(el)) {
                return;
            }
            // 若鼠标悬停在极小的行内公式图片或单字 span 上，自动向上吸附到最近的题目块或段落容器，方便一键选中整块内容
            const rect = el.getBoundingClientRect();
            if (rect.width < 180 && rect.height < 60 && el.parentElement && !isOwnPluginElement(el.parentElement)) {
                el = el.closest('[class*="quest"], [class*="timu"], [class*="paper"], article, section, li, p, div') || el.parentElement;
            }
            highlightElement(el);
        }

        function onClick(e) {
            if (isOwnPluginElement(e.target)) return;
            e.preventDefault();
            e.stopPropagation();
            const picked = hoveredTarget || e.target;
            cleanupPicker();
            openCleanView(picked, false);
        }

        function onKeyDown(e) {
            if (e.key === 'Escape') {
                cleanupPicker();
                showToast('已取消手动点选');
            } else if ((e.key === 'ArrowUp' || e.key === 'w' || e.key === 'W') && hoveredTarget && hoveredTarget.parentElement && hoveredTarget.parentElement !== document.documentElement) {
                e.preventDefault();
                highlightElement(hoveredTarget.parentElement);
                showToast('🔼 已扩大选区到上一层父容器（左键点击确认）', 2000);
            }
        }

        function cleanupPicker() {
            isPickerActive = false;
            if (hoveredTarget) {
                hoveredTarget.classList.remove(`${APP_PREFIX}-picker-highlight`);
                hoveredTarget = null;
            }
            document.removeEventListener('mousemove', onMouseMove, true);
            document.removeEventListener('click', onClick, true);
            document.removeEventListener('keydown', onKeyDown, true);
        }

        document.addEventListener('mousemove', onMouseMove, true);
        document.addEventListener('click', onClick, true);
        document.addEventListener('keydown', onKeyDown, true);
    }

    /**
     * 初始化悬浮按钮与纯净阅读弹窗
     */
    function initUI() {
        injectStyles();

        // 1. Toast 提示框
        toastEl = document.createElement('div');
        toastEl.id = `${APP_PREFIX}-toast`;
        document.body.appendChild(toastEl);

        // 2. 右下角悬浮按钮与展开面板
        const fabContainer = document.createElement('div');
        fabContainer.id = `${APP_PREFIX}-fab-container`;
        fabContainer.innerHTML = `
            <div id="${APP_PREFIX}-panel">
                <div class="${APP_PREFIX}-panel-title">
                    <span>📄 网页资料提取与导出</span>
                </div>
                <button class="${APP_PREFIX}-menu-item" data-action="clean-view">
                    <span>📖</span> <span>纯净阅读模式</span>
                </button>
                <button class="${APP_PREFIX}-menu-item" data-action="deep-extract">
                    <span>🔍</span> <span>深度提取 (含隐藏/伪元素)</span>
                </button>
                <button class="${APP_PREFIX}-menu-item" data-action="pick-area">
                    <span>🎯</span> <span>手动点选提取区域</span>
                </button>
                <button class="${APP_PREFIX}-menu-item" data-action="export-pdf">
                    <span>📄</span> <span>下载 PDF</span>
                </button>
                <button class="${APP_PREFIX}-menu-item" data-action="export-word">
                    <span>📝</span> <span>下载 Word</span>
                </button>
                <button class="${APP_PREFIX}-menu-item" data-action="system-print">
                    <span>🖨️</span> <span>系统打印 / 另存PDF</span>
                </button>
            </div>
            <button id="${APP_PREFIX}-fab-btn" title="提取网页正文并导出 PDF/Word">
                <span>📥 提取/导出</span>
            </button>
        `;
        document.body.appendChild(fabContainer);

        // 3. 纯净阅读模式全屏视图
        overlayEl = document.createElement('div');
        overlayEl.id = `${APP_PREFIX}-clean-overlay`;
        overlayEl.innerHTML = `
            <div id="${APP_PREFIX}-toolbar">
                <div class="${APP_PREFIX}-toolbar-left">
                    <span class="${APP_PREFIX}-toolbar-title">📖 纯净阅读与打印视图</span>
                    <span class="${APP_PREFIX}-toolbar-badge">已过滤侧边栏、广告与无关导航</span>
                </div>
                <div class="${APP_PREFIX}-toolbar-actions">
                    <button class="${APP_PREFIX}-btn" data-action="deep-extract-modal" title="遍历提取所有文字、图片、CSS隐藏节点(display:none/visibility:hidden)及::after伪元素">
                        🔍 深度提取(含隐藏/伪元素)
                    </button>
                    <button class="${APP_PREFIX}-btn" data-action="expand-parent" title="当前识别范围太小时，点击向上扩大一层父容器">
                        🔼 扩大范围
                    </button>
                    <button class="${APP_PREFIX}-btn" data-action="canvas-only" title="针对在线文档阅读器，仅提取已渲染的页面画布">
                        🖼️ 仅提画布页
                    </button>
                    <button class="${APP_PREFIX}-btn" data-action="re-pick" title="重新用鼠标手动框选网页区域">
                        🎯 手动选区
                    </button>
                    <button class="${APP_PREFIX}-btn" data-action="toggle-edit" title="点击后可直接增删修改正文文字">
                        ✏️ 微调编辑
                    </button>
                    <button class="${APP_PREFIX}-btn ${APP_PREFIX}-btn-primary" data-action="modal-pdf">
                        📄 导出 PDF
                    </button>
                    <button class="${APP_PREFIX}-btn ${APP_PREFIX}-btn-primary" data-action="modal-word">
                        📝 导出 Word
                    </button>
                    <button class="${APP_PREFIX}-btn" data-action="modal-print">
                        🖨️ 系统打印 (高清PDF)
                    </button>
                    <button class="${APP_PREFIX}-btn ${APP_PREFIX}-btn-danger" data-action="close-modal">
                        ✕ 关闭
                    </button>
                </div>
            </div>
            <div id="${APP_PREFIX}-scroll-area">
                <div id="${APP_PREFIX}-paper" spellcheck="false"></div>
            </div>
        `;
        document.body.appendChild(overlayEl);
        paperEl = overlayEl.querySelector(`#${APP_PREFIX}-paper`);

        // 4. 事件绑定
        const fabBtn = fabContainer.querySelector(`#${APP_PREFIX}-fab-btn`);
        const panel = fabContainer.querySelector(`#${APP_PREFIX}-panel`);

        fabBtn.addEventListener('click', (e) => {
            e.stopPropagation();
            panel.classList.toggle('open');
        });

        document.addEventListener('click', (e) => {
            if (!fabContainer.contains(e.target)) {
                panel.classList.remove('open');
            }
        });

        // 悬浮菜单点击事件
        panel.addEventListener('click', async (e) => {
            const btn = e.target.closest(`.${APP_PREFIX}-menu-item`);
            if (!btn) return;
            const action = btn.getAttribute('data-action');
            panel.classList.remove('open');

            // 若纯净视图已经打开且有内容，直接复用当前视图导出，不再重复识别覆盖用户已选/已编辑内容
            const isAlreadyOpen = overlayEl.classList.contains('active') && paperEl.innerHTML.trim().length > 0;

            if (action === 'clean-view') {
                await openCleanView();
            } else if (action === 'deep-extract') {
                await openDeepExtractView();
                showToast('🔍 已深度提取页面文字、图片、CSS隐藏节点、伪元素及内嵌 iframe 文档');
            } else if (action === 'pick-area') {
                if (overlayEl.classList.contains('active')) closeCleanView();
                startAreaPicker();
            } else if (action === 'export-pdf') {
                const container = isAlreadyOpen ? paperEl : await openCleanView();
                await exportToPDF(container);
            } else if (action === 'export-word') {
                const container = isAlreadyOpen ? paperEl : await openCleanView();
                await exportToWord(container);
            } else if (action === 'system-print') {
                if (!isAlreadyOpen) await openCleanView();
                triggerSystemPrint(paperEl);
            }
        });

        // 纯净阅读视图顶部工具栏事件
        overlayEl.querySelector(`#${APP_PREFIX}-toolbar`).addEventListener('click', async (e) => {
            const btn = e.target.closest('button[data-action]');
            if (!btn) return;
            const action = btn.getAttribute('data-action');

            if (action === 'close-modal') {
                closeCleanView();
            } else if (action === 'deep-extract-modal') {
                await openDeepExtractView(currentSourceNode || document.body);
                showToast('🔍 已深度提取当前区域所有隐藏节点、伪元素与内嵌 iframe 文档');
            } else if (action === 'expand-parent') {
                if (currentSourceNode && currentSourceNode.parentElement && currentSourceNode.parentElement !== document.documentElement) {
                    await openCleanView(currentSourceNode.parentElement);
                    showToast('🔼 已向上扩大一层识别范围');
                } else {
                    showToast('已到达顶层容器');
                }
            } else if (action === 'canvas-only') {
                const canvases = findDocumentPageCanvases(document.body);
                if (canvases.length > 0) {
                    await openCleanView(document.body, true);
                    showToast(`🖼️ 已提取页面中 ${canvases.length} 个已渲染的文档画布页`);
                } else {
                    showToast('⚠️ 当前页面未检测到已渲染的大尺寸文档 Canvas 画布');
                }
            } else if (action === 'toggle-edit') {
                const isEditable = paperEl.getAttribute('contenteditable') === 'true';
                paperEl.setAttribute('contenteditable', String(!isEditable));
                btn.textContent = !isEditable ? '✅ 完成编辑' : '✏️ 微调编辑';
                showToast(!isEditable ? '已开启编辑模式：可直接点击正文删改内容' : '已锁定正文内容');
            } else if (action === 're-pick') {
                closeCleanView();
                startAreaPicker();
            } else if (action === 'modal-pdf') {
                await exportToPDF(paperEl);
            } else if (action === 'modal-word') {
                await exportToWord(paperEl);
            } else if (action === 'modal-print') {
                triggerSystemPrint(paperEl);
            }
        });

        // 快捷键支持：Alt + R 快速开启/关闭纯净阅读模式
        document.addEventListener('keydown', (e) => {
            if (e.altKey && (e.key === 'r' || e.key === 'R')) {
                e.preventDefault();
                if (overlayEl.classList.contains('active')) {
                    closeCleanView();
                } else {
                    openCleanView();
                }
            }
        });

        // 监听滚动自动缓存已渲染的文档画布页（防止虚拟滚动清空已读页）
        let scrollCacheTimer = null;
        window.addEventListener('scroll', () => {
            clearTimeout(scrollCacheTimer);
            scrollCacheTimer = setTimeout(scanAndCacheVisibleCanvases, 300);
        }, { passive: true });
        setTimeout(scanAndCacheVisibleCanvases, 1000);

        // 支持 URL 参数 ?cre_auto=clean 自动进入纯净阅读与打印就绪状态
        if (window.location.search.includes('cre_auto=clean')) {
            openCleanView();
            document.body.classList.add(`${APP_PREFIX}-printing`);
        }
    }

    // 启动初始化
    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', initUI);
    } else {
        initUI();
    }
})();
