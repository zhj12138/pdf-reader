import markdown
from pymdownx import superfences, magiclink, caret, betterem, mark, tasklist, tilde, inlinehilite, critic, highlight,\
    _bypassnorm
import mdx_math

extensions = [
    'markdown.extensions.abbr',
    'markdown.extensions.attr_list',
    'markdown.extensions.def_list',
    'markdown.extensions.fenced_code',
    'markdown.extensions.footnotes',
    'markdown.extensions.md_in_html',
    'markdown.extensions.tables',
    'markdown.extensions.codehilite',
    'markdown.extensions.legacy_em',
    'markdown.extensions.toc',
    'markdown.extensions.wikilinks',
    'markdown.extensions.admonition',
    'markdown.extensions.legacy_attrs',
    'markdown.extensions.meta',
    'markdown.extensions.nl2br',
    'markdown.extensions.sane_lists',
    'markdown.extensions.smarty',
    'mdx_math',  # KaTeX数学公式，$E=mc^2$和$$E=mc^2$$
    # 'markdown_checklist.extension',  # checklist，- [ ]和- [x]
    'pymdownx.magiclink',  # 自动转超链接，
    'pymdownx.caret',  # 上标下标，
    'pymdownx.superfences',  # 多种块功能允许嵌套，各种图表
    'pymdownx.betterem',  # 改善强调的处理(粗体和斜体)
    'pymdownx.mark',  # 亮色突出文本
    # 'pymdownx.highlight',  # 高亮显示代码
    'pymdownx.tasklist',  # 任务列表
    'pymdownx.tilde',  # 删除线
    # 'pymdownx.emoji',
    'pymdownx.inlinehilite',
    # 'pymdownx.arithmatex',
    'pymdownx.critic',
    # 'markdown_markup_emoji.markup_emoji',
]
Mycss = '''
pre { line-height: 125%; margin: 0; }
td.linenos pre { color: #000000; background-color: #f0f0f0; padding: 0 5px 0 5px; }
span.linenos { color: #000000; background-color: #f0f0f0; padding: 0 5px 0 5px; }
td.linenos pre.special { color: #000000; background-color: #ffffc0; padding: 0 5px 0 5px; }
span.linenos.special { color: #000000; background-color: #ffffc0; padding: 0 5px 0 5px; }
.highlight .hll { background-color: #ffffcc }
.highlight { background: #f8f8f8; }
.highlight .c { color: #408080; font-style: italic } /* Comment */
.highlight .err { border: 1px solid #FF0000 } /* Error */
.highlight .k { color: #008000; font-weight: bold } /* Keyword */
.highlight .o { color: #666666 } /* Operator */
.highlight .ch { color: #408080; font-style: italic } /* Comment.Hashbang */
.highlight .cm { color: #408080; font-style: italic } /* Comment.Multiline */
.highlight .cp { color: #BC7A00 } /* Comment.Preproc */
.highlight .cpf { color: #408080; font-style: italic } /* Comment.PreprocFile */
.highlight .c1 { color: #408080; font-style: italic } /* Comment.Single */
.highlight .cs { color: #408080; font-style: italic } /* Comment.Special */
.highlight .gd { color: #A00000 } /* Generic.Deleted */
.highlight .ge { font-style: italic } /* Generic.Emph */
.highlight .gr { color: #FF0000 } /* Generic.Error */
.highlight .gh { color: #000080; font-weight: bold } /* Generic.Heading */
.highlight .gi { color: #00A000 } /* Generic.Inserted */
.highlight .go { color: #888888 } /* Generic.Output */
.highlight .gp { color: #000080; font-weight: bold } /* Generic.Prompt */
.highlight .gs { font-weight: bold } /* Generic.Strong */
.highlight .gu { color: #800080; font-weight: bold } /* Generic.Subheading */
.highlight .gt { color: #0044DD } /* Generic.Traceback */
.highlight .kc { color: #008000; font-weight: bold } /* Keyword.Constant */
.highlight .kd { color: #008000; font-weight: bold } /* Keyword.Declaration */
.highlight .kn { color: #008000; font-weight: bold } /* Keyword.Namespace */
.highlight .kp { color: #008000 } /* Keyword.Pseudo */
.highlight .kr { color: #008000; font-weight: bold } /* Keyword.Reserved */
.highlight .kt { color: #B00040 } /* Keyword.Type */
.highlight .m { color: #666666 } /* Literal.Number */
.highlight .s { color: #BA2121 } /* Literal.String */
.highlight .na { color: #7D9029 } /* Name.Attribute */
.highlight .nb { color: #008000 } /* Name.Builtin */
.highlight .nc { color: #0000FF; font-weight: bold } /* Name.Class */
.highlight .no { color: #880000 } /* Name.Constant */
.highlight .nd { color: #AA22FF } /* Name.Decorator */
.highlight .ni { color: #999999; font-weight: bold } /* Name.Entity */
.highlight .ne { color: #D2413A; font-weight: bold } /* Name.Exception */
.highlight .nf { color: #0000FF } /* Name.Function */
.highlight .nl { color: #A0A000 } /* Name.Label */
.highlight .nn { color: #0000FF; font-weight: bold } /* Name.Namespace */
.highlight .nt { color: #008000; font-weight: bold } /* Name.Tag */
.highlight .nv { color: #19177C } /* Name.Variable */
.highlight .ow { color: #AA22FF; font-weight: bold } /* Operator.Word */
.highlight .w { color: #bbbbbb } /* Text.Whitespace */
.highlight .mb { color: #666666 } /* Literal.Number.Bin */
.highlight .mf { color: #666666 } /* Literal.Number.Float */
.highlight .mh { color: #666666 } /* Literal.Number.Hex */
.highlight .mi { color: #666666 } /* Literal.Number.Integer */
.highlight .mo { color: #666666 } /* Literal.Number.Oct */
.highlight .sa { color: #BA2121 } /* Literal.String.Affix */
.highlight .sb { color: #BA2121 } /* Literal.String.Backtick */
.highlight .sc { color: #BA2121 } /* Literal.String.Char */
.highlight .dl { color: #BA2121 } /* Literal.String.Delimiter */
.highlight .sd { color: #BA2121; font-style: italic } /* Literal.String.Doc */
.highlight .s2 { color: #BA2121 } /* Literal.String.Double */
.highlight .se { color: #BB6622; font-weight: bold } /* Literal.String.Escape */
.highlight .sh { color: #BA2121 } /* Literal.String.Heredoc */
.highlight .si { color: #BB6688; font-weight: bold } /* Literal.String.Interpol */
.highlight .sx { color: #008000 } /* Literal.String.Other */
.highlight .sr { color: #BB6688 } /* Literal.String.Regex */
.highlight .s1 { color: #BA2121 } /* Literal.String.Single */
.highlight .ss { color: #19177C } /* Literal.String.Symbol */
.highlight .bp { color: #008000 } /* Name.Builtin.Pseudo */
.highlight .fm { color: #0000FF } /* Name.Function.Magic */
.highlight .vc { color: #19177C } /* Name.Variable.Class */
.highlight .vg { color: #19177C } /* Name.Variable.Global */
.highlight .vi { color: #19177C } /* Name.Variable.Instance */
.highlight .vm { color: #19177C } /* Name.Variable.Magic */
.highlight .il { color: #666666 } /* Literal.Number.Integer.Long */
[data-linenos]:before {
  content: attr(data-linenos);
}
.markdown-body .task-list-item {
  list-style-type: none !important;
}

.markdown-body .task-list-item input[type="checkbox"] {
  margin: 0 4px 0.25em -20px;
  vertical-align: middle;
}
.markdown-body .octicon {
  display: inline-block;
  fill: currentColor;
  vertical-align: text-bottom;
}

.markdown-body .anchor {
  float: left;
  line-height: 1;
  margin-left: -20px;
  padding-right: 4px;
}

.markdown-body .anchor:focus {
  outline: none;
}

.markdown-body h1 .octicon-link,
.markdown-body h2 .octicon-link,
.markdown-body h3 .octicon-link,
.markdown-body h4 .octicon-link,
.markdown-body h5 .octicon-link,
.markdown-body h6 .octicon-link {
  color: #1b1f23;
  vertical-align: middle;
  visibility: hidden;
}

.markdown-body h1:hover .anchor,
.markdown-body h2:hover .anchor,
.markdown-body h3:hover .anchor,
.markdown-body h4:hover .anchor,
.markdown-body h5:hover .anchor,
.markdown-body h6:hover .anchor {
  text-decoration: none;
}

.markdown-body h1:hover .anchor .octicon-link,
.markdown-body h2:hover .anchor .octicon-link,
.markdown-body h3:hover .anchor .octicon-link,
.markdown-body h4:hover .anchor .octicon-link,
.markdown-body h5:hover .anchor .octicon-link,
.markdown-body h6:hover .anchor .octicon-link {
  visibility: visible;
}

.markdown-body h1:hover .anchor .octicon-link:before,
.markdown-body h2:hover .anchor .octicon-link:before,
.markdown-body h3:hover .anchor .octicon-link:before,
.markdown-body h4:hover .anchor .octicon-link:before,
.markdown-body h5:hover .anchor .octicon-link:before,
.markdown-body h6:hover .anchor .octicon-link:before {
  width: 16px;
  height: 16px;
  content: ' ';
  display: inline-block;
  background-image: url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 16 16' version='1.1' width='16' height='16' aria-hidden='true'%3E%3Cpath fill-rule='evenodd' d='M4 9h1v1H4c-1.5 0-3-1.69-3-3.5S2.55 3 4 3h4c1.45 0 3 1.69 3 3.5 0 1.41-.91 2.72-2 3.25V8.59c.58-.45 1-1.27 1-2.09C10 5.22 8.98 4 8 4H4c-.98 0-2 1.22-2 2.5S3 9 4 9zm9-3h-1v1h1c1 0 2 1.22 2 2.5S13.98 12 13 12H9c-.98 0-2-1.22-2-2.5 0-.83.42-1.64 1-2.09V6.25c-1.09.53-2 1.84-2 3.25C6 11.31 7.55 13 9 13h4c1.45 0 3-1.69 3-3.5S14.5 6 13 6z'%3E%3C/path%3E%3C/svg%3E");
}.markdown-body {
  -ms-text-size-adjust: 100%;
  -webkit-text-size-adjust: 100%;
  line-height: 1.5;
  color: #24292e;
  font-family: -apple-system,BlinkMacSystemFont,Segoe UI,Helvetica,Arial,sans-serif,Apple Color Emoji,Segoe UI Emoji;
  font-size: 16px;
  line-height: 1.5;
  word-wrap: break-word;
}

.markdown-body details {
  display: block;
}

.markdown-body summary {
  display: list-item;
}

.markdown-body a {
  background-color: initial;
}

.markdown-body a:active,
.markdown-body a:hover {
  outline-width: 0;
}

.markdown-body strong {
  font-weight: inherit;
  font-weight: bolder;
}

.markdown-body h1 {
  font-size: 2em;
  margin: .67em 0;
}

.markdown-body img {
  border-style: none;
}

.markdown-body code,
.markdown-body kbd,
.markdown-body pre {
  font-family: monospace,monospace;
  font-size: 1em;
}

.markdown-body hr {
  box-sizing: initial;
  height: 0;
  overflow: visible;
}

.markdown-body input {
  font: inherit;
  margin: 0;
}

.markdown-body input {
  overflow: visible;
}

.markdown-body [type=checkbox] {
  box-sizing: border-box;
  padding: 0;
}

.markdown-body * {
  box-sizing: border-box;
}

.markdown-body input {
  font-family: inherit;
  font-size: inherit;
  line-height: inherit;
}

.markdown-body a {
  color: #0366d6;
  text-decoration: none;
}

.markdown-body a:hover {
  text-decoration: underline;
}

.markdown-body strong {
  font-weight: 600;
}

.markdown-body hr {
  height: 0;
  margin: 15px 0;
  overflow: hidden;
  background: transparent;
  border: 0;
  border-bottom: 1px solid #dfe2e5;
}

.markdown-body hr:after,
.markdown-body hr:before {
  display: table;
  content: "";
}

.markdown-body hr:after {
  clear: both;
}

.markdown-body table {
  border-spacing: 0;
  border-collapse: collapse;
}

.markdown-body td,
.markdown-body th {
  padding: 0;
}

.markdown-body details summary {
  cursor: pointer;
}

.markdown-body kbd {
  display: inline-block;
  padding: 3px 5px;
  font: 11px SFMono-Regular,Consolas,Liberation Mono,Menlo,monospace;
  line-height: 10px;
  color: #444d56;
  vertical-align: middle;
  background-color: #fafbfc;
  border: 1px solid #d1d5da;
  border-radius: 3px;
  box-shadow: inset 0 -1px 0 #d1d5da;
}

.markdown-body h1,
.markdown-body h2,
.markdown-body h3,
.markdown-body h4,
.markdown-body h5,
.markdown-body h6 {
  margin-top: 0;
  margin-bottom: 0;
}

.markdown-body h1 {
  font-size: 32px;
}

.markdown-body h1,
.markdown-body h2 {
  font-weight: 600;
}

.markdown-body h2 {
  font-size: 24px;
}

.markdown-body h3 {
  font-size: 20px;
}

.markdown-body h3,
.markdown-body h4 {
  font-weight: 600;
}

.markdown-body h4 {
  font-size: 16px;
}

.markdown-body h5 {
  font-size: 14px;
}

.markdown-body h5,
.markdown-body h6 {
  font-weight: 600;
}

.markdown-body h6 {
  font-size: 12px;
}

.markdown-body p {
  margin-top: 0;
  margin-bottom: 10px;
}

.markdown-body blockquote {
  margin: 0;
}

.markdown-body ol,
.markdown-body ul {
  padding-left: 0;
  margin-top: 0;
  margin-bottom: 0;
}

.markdown-body ol ol,
.markdown-body ul ol {
  list-style-type: lower-roman;
}

.markdown-body ol ol ol,
.markdown-body ol ul ol,
.markdown-body ul ol ol,
.markdown-body ul ul ol {
  list-style-type: lower-alpha;
}

.markdown-body dd {
  margin-left: 0;
}

.markdown-body code,
.markdown-body pre {
  font-family: SFMono-Regular,Consolas,Liberation Mono,Menlo,monospace;
  font-size: 12px;
}

.markdown-body pre {
  margin-top: 0;
  margin-bottom: 0;
}

.markdown-body input::-webkit-inner-spin-button,
.markdown-body input::-webkit-outer-spin-button {
  margin: 0;
  -webkit-appearance: none;
  appearance: none;
}

.markdown-body :checked+.radio-label {
  position: relative;
  z-index: 1;
  border-color: #0366d6;
}

.markdown-body .border {
  border: 1px solid #e1e4e8!important;
}

.markdown-body .border-0 {
  border: 0!important;
}

.markdown-body .border-bottom {
  border-bottom: 1px solid #e1e4e8!important;
}

.markdown-body .rounded-1 {
  border-radius: 3px!important;
}

.markdown-body .bg-white {
  background-color: #fff!important;
}

.markdown-body .bg-gray-light {
  background-color: #fafbfc!important;
}

.markdown-body .text-gray-light {
  color: #6a737d!important;
}

.markdown-body .mb-0 {
  margin-bottom: 0!important;
}

.markdown-body .my-2 {
  margin-top: 8px!important;
  margin-bottom: 8px!important;
}

.markdown-body .pl-0 {
  padding-left: 0!important;
}

.markdown-body .py-0 {
  padding-top: 0!important;
  padding-bottom: 0!important;
}

.markdown-body .pl-1 {
  padding-left: 4px!important;
}

.markdown-body .pl-2 {
  padding-left: 8px!important;
}

.markdown-body .py-2 {
  padding-top: 8px!important;
  padding-bottom: 8px!important;
}

.markdown-body .pl-3,
.markdown-body .px-3 {
  padding-left: 16px!important;
}

.markdown-body .px-3 {
  padding-right: 16px!important;
}

.markdown-body .pl-4 {
  padding-left: 24px!important;
}

.markdown-body .pl-5 {
  padding-left: 32px!important;
}

.markdown-body .pl-6 {
  padding-left: 40px!important;
}

.markdown-body .f6 {
  font-size: 12px!important;
}

.markdown-body .lh-condensed {
  line-height: 1.25!important;
}

.markdown-body .text-bold {
  font-weight: 600!important;
}

.markdown-body .pl-c {
  color: #6a737d;
}

.markdown-body .pl-c1,
.markdown-body .pl-s .pl-v {
  color: #005cc5;
}

.markdown-body .pl-e,
.markdown-body .pl-en {
  color: #6f42c1;
}

.markdown-body .pl-s .pl-s1,
.markdown-body .pl-smi {
  color: #24292e;
}

.markdown-body .pl-ent {
  color: #22863a;
}

.markdown-body .pl-k {
  color: #d73a49;
}

.markdown-body .pl-pds,
.markdown-body .pl-s,
.markdown-body .pl-s .pl-pse .pl-s1,
.markdown-body .pl-sr,
.markdown-body .pl-sr .pl-cce,
.markdown-body .pl-sr .pl-sra,
.markdown-body .pl-sr .pl-sre {
  color: #032f62;
}

.markdown-body .pl-smw,
.markdown-body .pl-v {
  color: #e36209;
}

.markdown-body .pl-bu {
  color: #b31d28;
}

.markdown-body .pl-ii {
  color: #fafbfc;
  background-color: #b31d28;
}

.markdown-body .pl-c2 {
  color: #fafbfc;
  background-color: #d73a49;
}

.markdown-body .pl-c2:before {
  content: "^M";
}

.markdown-body .pl-sr .pl-cce {
  font-weight: 700;
  color: #22863a;
}

.markdown-body .pl-ml {
  color: #735c0f;
}

.markdown-body .pl-mh,
.markdown-body .pl-mh .pl-en,
.markdown-body .pl-ms {
  font-weight: 700;
  color: #005cc5;
}

.markdown-body .pl-mi {
  font-style: italic;
  color: #24292e;
}

.markdown-body .pl-mb {
  font-weight: 700;
  color: #24292e;
}

.markdown-body .pl-md {
  color: #b31d28;
  background-color: #ffeef0;
}

.markdown-body .pl-mi1 {
  color: #22863a;
  background-color: #f0fff4;
}

.markdown-body .pl-mc {
  color: #e36209;
  background-color: #ffebda;
}

.markdown-body .pl-mi2 {
  color: #f6f8fa;
  background-color: #005cc5;
}

.markdown-body .pl-mdr {
  font-weight: 700;
  color: #6f42c1;
}

.markdown-body .pl-ba {
  color: #586069;
}

.markdown-body .pl-sg {
  color: #959da5;
}

.markdown-body .pl-corl {
  text-decoration: underline;
  color: #032f62;
}

.markdown-body .mb-0 {
  margin-bottom: 0!important;
}

.markdown-body .my-2 {
  margin-bottom: 8px!important;
}

.markdown-body .my-2 {
  margin-top: 8px!important;
}

.markdown-body .pl-0 {
  padding-left: 0!important;
}

.markdown-body .py-0 {
  padding-top: 0!important;
  padding-bottom: 0!important;
}

.markdown-body .pl-1 {
  padding-left: 4px!important;
}

.markdown-body .pl-2 {
  padding-left: 8px!important;
}

.markdown-body .py-2 {
  padding-top: 8px!important;
  padding-bottom: 8px!important;
}

.markdown-body .pl-3 {
  padding-left: 16px!important;
}

.markdown-body .pl-4 {
  padding-left: 24px!important;
}

.markdown-body .pl-5 {
  padding-left: 32px!important;
}

.markdown-body .pl-6 {
  padding-left: 40px!important;
}

.markdown-body .pl-7 {
  padding-left: 48px!important;
}

.markdown-body .pl-8 {
  padding-left: 64px!important;
}

.markdown-body .pl-9 {
  padding-left: 80px!important;
}

.markdown-body .pl-10 {
  padding-left: 96px!important;
}

.markdown-body .pl-11 {
  padding-left: 112px!important;
}

.markdown-body .pl-12 {
  padding-left: 128px!important;
}

.markdown-body hr {
  border-bottom-color: #eee;
}

.markdown-body kbd {
  display: inline-block;
  padding: 3px 5px;
  font: 11px SFMono-Regular,Consolas,Liberation Mono,Menlo,monospace;
  line-height: 10px;
  color: #444d56;
  vertical-align: middle;
  background-color: #fafbfc;
  border: 1px solid #d1d5da;
  border-radius: 3px;
  box-shadow: inset 0 -1px 0 #d1d5da;
}

.markdown-body:after,
.markdown-body:before {
  display: table;
  content: "";
}

.markdown-body:after {
  clear: both;
}

.markdown-body>:first-child {
  margin-top: 0!important;
}

.markdown-body>:last-child {
  margin-bottom: 0!important;
}

.markdown-body a:not([href]) {
  color: inherit;
  text-decoration: none;
}

.markdown-body blockquote,
.markdown-body details,
.markdown-body dl,
.markdown-body ol,
.markdown-body p,
.markdown-body pre,
.markdown-body table,
.markdown-body ul {
  margin-top: 0;
  margin-bottom: 16px;
}

.markdown-body hr {
  height: .25em;
  padding: 0;
  margin: 24px 0;
  background-color: #e1e4e8;
  border: 0;
}

.markdown-body blockquote {
  padding: 0 1em;
  color: #6a737d;
  border-left: .25em solid #dfe2e5;
}

.markdown-body blockquote>:first-child {
  margin-top: 0;
}

.markdown-body blockquote>:last-child {
  margin-bottom: 0;
}

.markdown-body h1,
.markdown-body h2,
.markdown-body h3,
.markdown-body h4,
.markdown-body h5,
.markdown-body h6 {
  margin-top: 24px;
  margin-bottom: 16px;
  font-weight: 600;
  line-height: 1.25;
}

.markdown-body h1 {
  font-size: 2em;
}

.markdown-body h1,
.markdown-body h2 {
  padding-bottom: .3em;
  border-bottom: 1px solid #eaecef;
}

.markdown-body h2 {
  font-size: 1.5em;
}

.markdown-body h3 {
  font-size: 1.25em;
}

.markdown-body h4 {
  font-size: 1em;
}

.markdown-body h5 {
  font-size: .875em;
}

.markdown-body h6 {
  font-size: .85em;
  color: #6a737d;
}

.markdown-body ol,
.markdown-body ul {
  padding-left: 2em;
}

.markdown-body ol ol,
.markdown-body ol ul,
.markdown-body ul ol,
.markdown-body ul ul {
  margin-top: 0;
  margin-bottom: 0;
}

.markdown-body li {
  word-wrap: break-all;
}

.markdown-body li>p {
  margin-top: 16px;
}

.markdown-body li+li {
  margin-top: .25em;
}

.markdown-body dl {
  padding: 0;
}

.markdown-body dl dt {
  padding: 0;
  margin-top: 16px;
  font-size: 1em;
  font-style: italic;
  font-weight: 600;
}

.markdown-body dl dd {
  padding: 0 16px;
  margin-bottom: 16px;
}

.markdown-body table {
  display: block;
  width: 100%;
  overflow: auto;
}

.markdown-body table th {
  font-weight: 600;
}

.markdown-body table td,
.markdown-body table th {
  padding: 6px 13px;
  border: 1px solid #dfe2e5;
}

.markdown-body table tr {
  background-color: #fff;
  border-top: 1px solid #c6cbd1;
}

.markdown-body table tr:nth-child(2n) {
  background-color: #f6f8fa;
}

.markdown-body img {
  max-width: 100%;
  box-sizing: initial;
  background-color: #fff;
}

.markdown-body img[align=right] {
  padding-left: 20px;
}

.markdown-body img[align=left] {
  padding-right: 20px;
}

.markdown-body code {
  padding: .2em .4em;
  margin: 0;
  font-size: 85%;
  background-color: rgba(27,31,35,.05);
  border-radius: 3px;
}

.markdown-body pre {
  word-wrap: normal;
}

.markdown-body pre>code {
  padding: 0;
  margin: 0;
  font-size: 100%;
  word-break: normal;
  white-space: pre;
  background: transparent;
  border: 0;
}

.markdown-body .highlight {
  margin-bottom: 16px;
}

.markdown-body .highlight pre {
  margin-bottom: 0;
  word-break: normal;
}

.markdown-body .highlight pre,
.markdown-body pre {
  padding: 16px;
  overflow: auto;
  font-size: 85%;
  line-height: 1.45;
  background-color: #f6f8fa;
  border-radius: 3px;
}

.markdown-body pre code {
  display: inline;
  max-width: auto;
  padding: 0;
  margin: 0;
  overflow: visible;
  line-height: inherit;
  word-wrap: normal;
  background-color: initial;
  border: 0;
}

.markdown-body .commit-tease-sha {
  display: inline-block;
  font-family: SFMono-Regular,Consolas,Liberation Mono,Menlo,monospace;
  font-size: 90%;
  color: #444d56;
}

.markdown-body .full-commit .btn-outline:not(:disabled):hover {
  color: #005cc5;
  border-color: #005cc5;
}

.markdown-body .blob-wrapper {
  overflow-x: auto;
  overflow-y: hidden;
}

.markdown-body .blob-wrapper-embedded {
  max-height: 240px;
  overflow-y: auto;
}

.markdown-body .blob-num {
  width: 1%;
  min-width: 50px;
  padding-right: 10px;
  padding-left: 10px;
  font-family: SFMono-Regular,Consolas,Liberation Mono,Menlo,monospace;
  font-size: 12px;
  line-height: 20px;
  color: rgba(27,31,35,.3);
  text-align: right;
  white-space: nowrap;
  vertical-align: top;
  cursor: pointer;
  -webkit-user-select: none;
  -moz-user-select: none;
  -ms-user-select: none;
  user-select: none;
}

.markdown-body .blob-num:hover {
  color: rgba(27,31,35,.6);
}

.markdown-body .blob-num:before {
  content: attr(data-line-number);
}

.markdown-body .blob-code {
  position: relative;
  padding-right: 10px;
  padding-left: 10px;
  line-height: 20px;
  vertical-align: top;
}

.markdown-body .blob-code-inner {
  overflow: visible;
  font-family: SFMono-Regular,Consolas,Liberation Mono,Menlo,monospace;
  font-size: 12px;
  color: #24292e;
  word-wrap: normal;
  white-space: pre;
}

.markdown-body .pl-token.active,
.markdown-body .pl-token:hover {
  cursor: pointer;
  background: #ffea7f;
}

.markdown-body .tab-size[data-tab-size="1"] {
  -moz-tab-size: 1;
  tab-size: 1;
}

.markdown-body .tab-size[data-tab-size="2"] {
  -moz-tab-size: 2;
  tab-size: 2;
}

.markdown-body .tab-size[data-tab-size="3"] {
  -moz-tab-size: 3;
  tab-size: 3;
}

.markdown-body .tab-size[data-tab-size="4"] {
  -moz-tab-size: 4;
  tab-size: 4;
}

.markdown-body .tab-size[data-tab-size="5"] {
  -moz-tab-size: 5;
  tab-size: 5;
}

.markdown-body .tab-size[data-tab-size="6"] {
  -moz-tab-size: 6;
  tab-size: 6;
}

.markdown-body .tab-size[data-tab-size="7"] {
  -moz-tab-size: 7;
  tab-size: 7;
}

.markdown-body .tab-size[data-tab-size="8"] {
  -moz-tab-size: 8;
  tab-size: 8;
}

.markdown-body .tab-size[data-tab-size="9"] {
  -moz-tab-size: 9;
  tab-size: 9;
}

.markdown-body .tab-size[data-tab-size="10"] {
  -moz-tab-size: 10;
  tab-size: 10;
}

.markdown-body .tab-size[data-tab-size="11"] {
  -moz-tab-size: 11;
  tab-size: 11;
}

.markdown-body .tab-size[data-tab-size="12"] {
  -moz-tab-size: 12;
  tab-size: 12;
}

.markdown-body .task-list-item {
  list-style-type: none;
}

.markdown-body .task-list-item+.task-list-item {
  margin-top: 3px;
}

.markdown-body .task-list-item input {
  margin: 0 .2em .25em -1.6em;
  vertical-align: middle;
}
'''
myhtml = '''
<!DOCTYPE html>
        <head>
            <meta charset="UTF-8">
            <meta name="viewport" content="width=device-width, initial-scale=1, minimal-ui">
            <title>Title</title>
            <style type="text/css">
                {}
            </style>
            <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/katex/dist/katex.min.css" crossorigin="anonymous">
            <script src="https://unpkg.com/mermaid@8.7.0/dist/mermaid.min.js"></script>
            <script src="https://cdn.jsdelivr.net/npm/katex/dist/katex.min.js" crossorigin="anonymous"></script>
            <script src="https://cdn.jsdelivr.net/npm/katex/dist/contrib/mathtex-script-type.min.js" defer></script>
        </head>
        <body>
            <article class="markdown-body">
                {}
            </article>
        </body>
    </html>
'''
extension_configs = {
    'mdx_math': {
        'enable_dollar_delimiter': True  # 允许单个$
    },
    'pymdownx.superfences': {
        "custom_fences": [
            {
                'name': 'mermaid',  # 开启流程图等图
                'class': 'mermaid',
                'format': superfences.fence_div_format
            }
        ]
    },
    'pymdownx.highlight': {
        'linenums': True,  # 显示行号
        'linenums_style': 'pymdownx-inline'  # 代码和行号分开
    },
    'pymdownx.tasklist': {
        'clickable_checkbox': True,  # 任务列表可点击
    }
}  # 扩展配置


def md2html(text):
    html = markdown.markdown(text, extensions=extensions, extension_configs=extension_configs)
    return myhtml.format(Mycss, html)
