// 한글(HWPX) 저장에 쓰는 빈 문서 틀(js/hwpx-template.js)을 만듭니다.
// 원본: python-hwpx 패키지의 hwpx/data/Skeleton.hwpx (Apache-2.0, Copyright 2025-2026 airmang)
// 실행: node tools/make_hwpx_template.js <Skeleton.hwpx 경로>
const fs = require('fs');
const path = require('path');
const JSZip = require('jszip');

(async () => {
  const src = process.argv[2];
  if (!src) throw new Error('Skeleton.hwpx 경로를 지정하세요.');
  const zip = await JSZip.loadAsync(fs.readFileSync(src));
  const keep = ['mimetype', 'version.xml', 'settings.xml', 'META-INF/container.xml', 'META-INF/container.rdf', 'META-INF/manifest.xml', 'Contents/content.hpf', 'Contents/header.xml', 'Contents/section0.xml'];
  const files = {};
  for (const name of keep) files[name] = await zip.file(name).async('string');
  // 미리보기 이미지는 넣지 않음(내용과 다른 그림이 보이지 않도록)
  files['META-INF/container.xml'] = files['META-INF/container.xml'].replace(/<ocf:rootfile full-path="Preview\/PrvImage\.png"[^>]*\/>/, '');
  const out = `/*
 * 한글(HWPX) 빈 문서 틀 — tools/make_hwpx_template.js 로 생성한 파일입니다. 직접 고치지 마세요.
 * 원본: python-hwpx (https://github.com/airmang/python-hwpx) hwpx/data/Skeleton.hwpx
 * Copyright 2025-2026 airmang. Licensed under the Apache License, Version 2.0.
 */
(function (root) {
  'use strict';
  const TG = (root.TG = root.TG || {});
  TG.HWPX_TEMPLATE = ${JSON.stringify(files, null, 1)};
  if (typeof module !== 'undefined' && module.exports) module.exports = TG;
})(typeof window !== 'undefined' ? window : globalThis);
`;
  fs.writeFileSync(path.join(__dirname, '..', 'js', 'hwpx-template.js'), out);
  console.log('js/hwpx-template.js 생성 완료');
})();
