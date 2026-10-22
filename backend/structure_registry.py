"""`backend/structure_registry.py` — SỔ ĐĂNG KÝ CẤU TRÚC LaTeX.

Đây là nơi **khai báo**, không phải nơi viết thuật toán.  Muốn dạy hệ thống một cấu
trúc mới thì chỉ cần thêm một dòng vào đúng tập hợp bên dưới — không phải sửa code.

Bốn mức xử lý (tier):

| Mức | Ý nghĩa | Kết quả trong Word |
|---|---|---|
| `NATIVE` | engine gốc đã dựng được | chữ / công thức / bảng **sửa được** |
| `PICTURE` | phải để TeX dựng rồi chụp lại | ảnh **vector**, giống bản gốc 100% |
| `TEXT` | lệnh trang trí, bỏ tên giữ nội dung | chữ thường |
| `IGNORE` | chỉ có nghĩa ở preamble | không in ra |

Lệnh/môi trường **không nằm trong sổ** sẽ được xử lý theo chính sách
(`backend/universal_fallback.py`): mặc định môi trường lạ → `PICTURE` (an toàn tuyệt
đối), lệnh lạ → nhường cho bộ tiền xử lý bóc vỏ.
"""
from __future__ import annotations

__all__ = [
    "NATIVE_ENVIRONMENTS", "PICTURE_ENVIRONMENTS", "TEXTUAL_ENVIRONMENTS", "IGNORE_ENVIRONMENTS",
    "NATIVE_COMMANDS", "PICTURE_COMMANDS", "TEXT_COMMANDS", "IGNORE_COMMANDS", "MATH_SYMBOLS",
    "PACKAGE_COMMANDS", "INLINE_PICTURE_COMMANDS",
    "VERBATIM_ENVIRONMENTS", "tier_of_environment", "tier_of_command", "registry_size",
]

# ─────────────────────────── MÔI TRƯỜNG ───────────────────────────

#: Môi trường engine gốc đã dựng được thành đối tượng Word sửa trực tiếp.
NATIVE_ENVIRONMENTS: set[str] = {
    # toán
    "equation", "equation*", "align", "align*", "alignat", "alignat*", "aligned", "alignedat",
    "gather", "gather*", "gathered", "multline", "multline*", "split", "flalign", "flalign*",
    "eqnarray", "eqnarray*", "array", "cases", "dcases", "rcases", "numcases", "subnumcases",
    "matrix", "pmatrix", "bmatrix", "Bmatrix", "vmatrix", "Vmatrix", "smallmatrix",
    "displaymath", "math", "subequations", "empheq",
    # bảng
    "tabular", "tabular*", "tabularx", "tabulary", "longtable", "ltabulary", "supertabular",
    "xltabular", "longtblr", "tblr", "talltblr", "booktabs", "NiceTabular", "NiceArray",
    # danh sách
    "itemize", "enumerate", "description", "list", "trivlist", "enumerate*", "itemize*",
    # văn bản
    "center", "flushleft", "flushright", "quote", "quotation", "verse", "abstract",
    "minipage", "figure", "figure*", "table", "table*", "wrapfigure", "wraptable",
    "multicols", "multicols*", "spacing", "samepage", "adjustwidth", "small", "footnotesize",
    # đề thi tiếng Việt (bộ ex_test / exam / ncctex)
    "ex", "ex_test", "bt", "baitap", "questions", "cauhoi", "dapan", "loigiai", "huongdan",
    "exam", "examexercise", "bonpa", "boncot", "haicot", "motcot", "choice", "choiceTF",
    # định lý
    "theorem", "lemma", "corollary", "proposition", "definition", "example", "remark",
    "proof", "dinhly", "dinhnghia", "vidu", "nhanxet", "chuy", "tinhchat", "hequa", "bode",
}

#: Môi trường **bắt buộc** để TeX dựng rồi nhúng ảnh vector (hình vẽ, đồ thị, sơ đồ…).
PICTURE_ENVIRONMENTS: set[str] = {
    # TikZ / PGF và họ hàng
    "tikzpicture", "scope", "pgfpicture", "pgfonlayer", "tikzcd", "tikzcd*",
    "axis", "semilogxaxis", "semilogyaxis", "loglogaxis", "polaraxis", "groupplot",
    "circuitikz", "forest", "dynkinDiagram", "ganttchart", "tikztimingtable",
    "smartdiagram", "chemfig", "schemescheme", "scheme", "venndiagram3sets",
    "feynman", "feynhand", "tikzfeynman", "bytefield", "neuralnetwork",
    # bảng biến thiên / bảng xét dấu
    "tkzTab", "tkztab", "tabvar", "tabular_bbt", "bbt", "variations", "tkzTabEnv",
    # vẽ thuần LaTeX
    "picture", "pspicture", "psmatrix", "asy", "asydef", "pgfplotsonlayer",
    # nhạc / cờ / hoá
    "music", "chessboard", "chessgame", "Chessboard", "xymatrix", "diagram",
    # khung màu có bo góc, bóng đổ — vẽ lại trong Word không bao giờ giống
    "tcolorbox", "tcblisting", "mdframed", "shadedbox", "framedbox",
    # đồ thị khác
    "pgfplotsgroupplot", "pgfgantt", "sequencediagram", "tikzsymbols",
}

#: Môi trường chỉ cần giữ nội dung chữ bên trong.
TEXTUAL_ENVIRONMENTS: set[str] = {
    "document", "sloppypar", "raggedright", "raggedleft", "em", "normalsize",
    "titlepage", "thebibliography", "appendix", "landscape", "otherlanguage",
}

#: Môi trường chỉ có nghĩa ở preamble / không in gì ra.
IGNORE_ENVIRONMENTS: set[str] = {"comment", "filecontents", "filecontents*", "NoHyper"}

#: Môi trường nguyên văn — tuyệt đối không được đụng vào nội dung.
VERBATIM_ENVIRONMENTS: set[str] = {
    "verbatim", "verbatim*", "Verbatim", "BVerbatim", "LVerbatim", "lstlisting",
    "minted", "alltt", "listing", "pycode", "sageblock", "asy",
}

# ─────────────────────────── LỆNH ───────────────────────────

#: Lệnh engine gốc đã dựng được (giữ nguyên để engine xử lý).
NATIVE_COMMANDS: set[str] = {
    # cấu trúc tài liệu
    "section", "subsection", "subsubsection", "paragraph", "subparagraph", "chapter", "part",
    "section*", "subsection*", "subsubsection*", "title", "author", "date", "maketitle",
    "item", "caption", "label", "ref", "pageref", "eqref", "cite", "footnote", "footnotetext",
    # định dạng chữ
    "textbf", "textit", "texttt", "textsf", "textrm", "textsc", "textsl", "emph", "underline",
    "uline", "uuline", "uwave", "sout", "xout", "textcolor", "colorbox", "fcolorbox", "hl",
    "tiny", "scriptsize", "footnotesize", "small", "normalsize", "large", "Large", "LARGE",
    "huge", "Huge", "bfseries", "itshape", "ttfamily", "rmfamily", "sffamily", "upshape",
    "mdseries", "scshape", "slshape", "normalfont", "textsuperscript", "textsubscript",
    # ngắt dòng / khoảng trắng
    "\\", "newline", "linebreak", "pagebreak", "newpage", "clearpage", "cleardoublepage",
    "par", "noindent", "indent", "hspace", "vspace", "hfill", "vfill", "quad", "qquad",
    "smallskip", "medskip", "bigskip", "dotfill", "hrulefill", "centering", "raggedright",
    "raggedleft", "tabcolsep", "arraystretch", "setlength", "addtolength",
    # toán — cơ bản
    "frac", "dfrac", "tfrac", "cfrac", "sqrt", "sum", "prod", "int", "iint", "iiint", "oint",
    "lim", "limsup", "liminf", "max", "min", "sup", "inf", "log", "ln", "lg", "exp", "sin",
    "cos", "tan", "cot", "sec", "csc", "arcsin", "arccos", "arctan", "sinh", "cosh", "tanh",
    "binom", "dbinom", "tbinom", "choose", "over", "atop", "overline", "underline", "widehat",
    "widetilde", "overrightarrow", "overleftarrow", "vec", "hat", "bar", "tilde", "dot", "ddot",
    "left", "right", "big", "Big", "bigg", "Bigg", "bigl", "bigr", "Bigl", "Bigr",
    "mathbb", "mathcal", "mathfrak", "mathbf", "mathit", "mathrm", "mathsf", "mathtt", "boldsymbol",
    "bm", "text", "textnormal", "mbox", "operatorname", "DeclareMathOperator", "substack",
    "stackrel", "overset", "underset", "xrightarrow", "xleftarrow", "xrightleftharpoons",
    "cancel", "bcancel", "xcancel", "cancelto", "phantom", "hphantom", "vphantom",
    "limits", "nolimits", "displaystyle", "textstyle", "scriptstyle", "scriptscriptstyle",
    # bảng
    "hline", "cline", "toprule", "midrule", "bottomrule", "cmidrule", "multicolumn", "multirow",
    "rowcolor", "cellcolor", "columncolor", "diagbox", "makecell", "thead", "tabularnewline",
    "arraybackslash", "extracolsep",
    # đề thi tiếng Việt
    "cau", "Cau", "choiceA", "choiceB", "choiceC", "choiceD", "choice", "choiceTF", "choicef",
    "bonpa", "boncot", "haicot", "motcot", "loigiai", "shortans", "immini", "True", "dapan",
    "hoanthanh", "khoanhdap", "dethi", "monthi",
}

#: Lệnh sinh ra hình vẽ → phải để TeX dựng.
PICTURE_COMMANDS: set[str] = {
    "tikz", "tikzset", "pgfplotstabletypeset", "chemfig", "schemestart", "ce", "cee",
    "includegraphics", "includesvg", "includepdf", "animategraphics",
    "tkzTabInit", "tkzTabLine", "tkzTabVar", "tkzTabVal", "tkzTabIma", "tkzInit", "tkzAxeXY",
    "tkzGrid", "tkzDrawX", "tkzDrawY", "tkzFct", "tkzDefPoint", "tkzDrawPolygon",
    "psset", "psline", "pspolygon", "psframe", "pscircle", "psplot", "psaxes",
    "xymatrix", "Qcircuit", "setchessboard", "chessboard", "newgame", "printchessboard",
    "smartdiagramset", "smartdiagram", "feynmandiagram", "pgfdeclareimage", "pgfuseimage",
    "barcode", "qrcode", "QRCode", "scalebox", "resizebox", "rotatebox", "reflectbox",
}

#: Lệnh trang trí — bỏ tên, giữ nội dung đối số cuối.
TEXT_COMMANDS: set[str] = {
    "mbox", "makebox", "framebox", "fbox", "raisebox", "parbox", "ensuremath",
    "textstyle", "protect", "relax", "mathsurround", "allowbreak", "nobreak",
}

#: Lệnh chỉ có nghĩa ở preamble / cấu hình — không in gì.
IGNORE_COMMANDS: set[str] = {
    "documentclass", "usepackage", "RequirePackage", "PassOptionsToPackage", "geometry",
    "pagestyle", "thispagestyle", "fancyhf", "fancyhead", "fancyfoot", "fancypagestyle",
    "headrulewidth", "footrulewidth", "renewcommand", "newcommand", "providecommand",
    "DeclareRobustCommand", "newenvironment", "renewenvironment", "newtheorem", "theoremstyle",
    "def", "let", "newif", "newcounter", "setcounter", "addtocounter", "stepcounter",
    "pgfplotsset", "tikzstyle", "definecolor", "colorlet", "captionsetup", "titleformat",
    "setmainfont", "setsansfont", "setmonofont", "newfontfamily", "babelprovide",
    "makeatletter", "makeatother", "AtBeginDocument", "AtEndDocument", "input", "include",
    "includeonly", "bibliography", "bibliographystyle", "printbibliography", "addbibresource",
    "tableofcontents", "listoffigures", "listoftables", "hypersetup", "sisetup", "lstset",
    "numberwithin", "allowdisplaybreaks", "setlist", "newlist", "setcolsep", "columnsep",
    "DeclarePairedDelimiter", "newcolumntype", "SetKwInOut", "graphicspath", "counterwithin",
}


# ─────────────────────── KÝ HIỆU TOÁN HỌC CHUẨN ───────────────────────
#: Toàn bộ ký hiệu toán của LaTeX/AMS — engine gốc dịch thẳng sang OMML.
MATH_SYMBOLS: set[str] = set("""
alpha beta gamma delta epsilon varepsilon zeta eta theta vartheta iota kappa lambda mu nu xi
pi varpi rho varrho sigma varsigma tau upsilon phi varphi chi psi omega
Gamma Delta Theta Lambda Xi Pi Sigma Upsilon Phi Psi Omega digamma varkappa
aleph beth gimel daleth hbar hslash imath jmath ell wp Re Im partial infty emptyset varnothing
nabla surd top bot angle measuredangle sphericalangle triangle forall exists nexists neg lnot
flat natural sharp clubsuit diamondsuit heartsuit spadesuit prime backprime circledS complement
eth mho Finv Game backepsilon
pm mp times div cdot ast star dagger ddagger amalg cap cup uplus sqcap sqcup vee wedge
setminus wr diamond bigtriangleup bigtriangledown triangleleft triangleright lhd rhd unlhd unrhd
oplus ominus otimes oslash odot bigcirc circ bullet dotplus ltimes rtimes leftthreetimes
rightthreetimes curlywedge curlyvee veebar barwedge doublebarwedge boxplus boxminus boxtimes
boxdot divideontimes centerdot intercal circledast circledcirc circleddash
leq le geq ge neq ne equiv sim simeq approx approxeq cong asymp doteq propto models perp mid
parallel nmid nparallel prec succ preceq succeq ll gg subset supset subseteq supseteq sqsubset
sqsupset sqsubseteq sqsupseteq in ni notin owns vdash dashv smile frown bowtie
nless ngtr nleq ngeq nleqslant ngeqslant lneq gneq lnsim gnsim precnsim succnsim nsubseteq
nsupseteq subsetneq supsetneq varsubsetneq varsupsetneq nsim ncong nvdash nvDash nVdash
leqslant geqslant eqslantless eqslantgtr lesssim gtrsim lessapprox gtrapprox approxeq lessdot
gtrdot lll ggg lessgtr gtrless lesseqgtr gtreqless lesseqqgtr gtreqqless doteqdot risingdotseq
fallingdotseq backsim backsimeq vDash Vdash Vvdash because therefore between pitchfork
varpropto blacktriangleleft blacktriangleright blacktriangle blacktriangledown
leftarrow rightarrow gets to leftrightarrow Leftarrow Rightarrow Leftrightarrow longleftarrow
longrightarrow longleftrightarrow Longleftarrow Longrightarrow Longleftrightarrow iff implies
impliedby mapsto longmapsto hookleftarrow hookrightarrow leftharpoonup leftharpoondown
rightharpoonup rightharpoondown rightleftharpoons leftrightharpoons uparrow downarrow updownarrow
Uparrow Downarrow Updownarrow nearrow searrow swarrow nwarrow leadsto rightsquigarrow
leftrightsquigarrow twoheadrightarrow twoheadleftarrow rightarrowtail leftarrowtail
curvearrowleft curvearrowright circlearrowleft circlearrowright upharpoonright upharpoonleft
downharpoonright downharpoonleft multimap nrightarrow nleftarrow nRightarrow nLeftarrow
nleftrightarrow nLeftrightarrow dashrightarrow dashleftarrow
sum prod coprod bigcup bigcap bigsqcup bigvee bigwedge bigodot bigotimes bigoplus biguplus
int iint iiint iiiint idotsint oint oiint oiiint intop smallint
arccos arcsin arctan arg cos cosh cot coth csc deg det dim exp gcd hom inf ker lg lim liminf
limsup ln log max min Pr sec sin sinh sup tan tanh
ldots cdots vdots ddots dots dotsb dotsc dotsi dotsm dotso hdots
langle rangle lfloor rfloor lceil rceil lvert rvert lVert rVert vert Vert backslash
lbrace rbrace lbrack rbrack llbracket rrbracket ulcorner urcorner llcorner lrcorner
quad qquad thinspace negthinspace medspace thickspace negmedspace negthickspace enspace
colon mathbin mathrel mathop mathord mathopen mathclose mathpunct mathinner
nolimits limits displaystyle textstyle scriptstyle scriptscriptstyle
degree celsius percent permil micro ohm angstrom pounds euro textdegree
dfrac tfrac genfrac sfrac nicefrac
acute grave breve check ring tilde hat dot ddot dddot ddddot vec bar mathring
overbrace underbrace overparen underparen overbracket underbracket overgroup undergroup
xleftrightarrow xhookrightarrow xmapsto xRightarrow xLeftarrow xLongrightarrow
pmod bmod mod pod sqrt root uproot leftroot
mathstrut smash llap rlap clap
circledR maltese checkmark square blacksquare lozenge blacklozenge bigstar
""".split())

#: Ký hiệu toán cũng là "dựng được bằng engine gốc".
NATIVE_COMMANDS |= MATH_SYMBOLS

#: Lệnh của các gói thông dụng mà engine gốc dịch thẳng sang chữ/công thức.
PACKAGE_COMMANDS: set[str] = set("""
SI si num qty unit ang numlist numrange SIlist SIrange numproduct complexnum
href url nolinkurl autoref nameref cref Cref crefrange hyperref hyperlink hypertarget
setlist newlist setlistdepth restartlist resume
footnotemark footnotetext thanks marginpar marginnote sidenote
newtcolorbox tcbset tcblower
rule hrule vrule strut smash mathstrut
textcircled circled textendash textemdash textquotedblleft textquotedblright
textless textgreater textbackslash textasciitilde textasciicircum textbar
ldots dots textellipsis guillemotleft guillemotright
and affil institute email orcid
captionof subcaption subfloat sidecaption
todo missingfigure listoftodos
qedhere blacksquare lipsum blindtext
""".split())
NATIVE_COMMANDS |= PACKAGE_COMMANDS

#: Lệnh vẽ/hoá học nằm GIỮA DÒNG — chụp ảnh inline, căn đúng đường cơ sở.
INLINE_PICTURE_COMMANDS: set[str] = {"ce", "cee", "chemfig", "qrcode", "QRCode", "tikzmark"}

_TIER_ENV = [(NATIVE_ENVIRONMENTS, "NATIVE"), (PICTURE_ENVIRONMENTS, "PICTURE"),
             (TEXTUAL_ENVIRONMENTS, "TEXT"), (IGNORE_ENVIRONMENTS, "IGNORE"),
             (VERBATIM_ENVIRONMENTS, "VERBATIM")]
_TIER_CMD = [(NATIVE_COMMANDS, "NATIVE"), (PICTURE_COMMANDS, "PICTURE"),
             (TEXT_COMMANDS, "TEXT"), (IGNORE_COMMANDS, "IGNORE")]


def tier_of_environment(name: str) -> str | None:
    """`NATIVE` | `PICTURE` | `TEXT` | `IGNORE` | `VERBATIM`, hoặc `None` nếu chưa biết."""
    bare = name.rstrip("*")
    for table, tier in _TIER_ENV:
        if name in table or bare in table:
            return tier
    return None


def tier_of_command(name: str) -> str | None:
    """`NATIVE` | `PICTURE` | `TEXT` | `IGNORE`, hoặc `None` nếu chưa biết."""
    for table, tier in _TIER_CMD:
        if name in table:
            return tier
    return None


def registry_size() -> dict[str, int]:
    return {
        "environments": len(NATIVE_ENVIRONMENTS | PICTURE_ENVIRONMENTS | TEXTUAL_ENVIRONMENTS
                            | IGNORE_ENVIRONMENTS | VERBATIM_ENVIRONMENTS),
        "commands": len(NATIVE_COMMANDS | PICTURE_COMMANDS | TEXT_COMMANDS | IGNORE_COMMANDS),
        "native_environments": len(NATIVE_ENVIRONMENTS),
        "picture_environments": len(PICTURE_ENVIRONMENTS),
        "native_commands": len(NATIVE_COMMANDS),
        "picture_commands": len(PICTURE_COMMANDS),
    }
