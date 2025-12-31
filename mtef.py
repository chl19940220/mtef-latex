from io import BytesIO
from .ole_util.helper import Helper
from .ole_util.ole import Ole
from .record import MtLine, MtChar, MtTmpl, MtPile, MtMatrix, MtEmbellRd, MtfontStyleDef, MtSize, MtfontDef, \
    MtColorDefIndex, MtColorDef, MtEqnPrefs, RecordType, OptionType, CharTypeface, SelectorType, EmbellType, MtAST, \
    RecordTypeV3, TagTypeV3, MTCharV3, EmbellTypeV3, SelectorTypeV3
from .chars import Chars, SpecialChar
from thesis_guru.utils.logger import get_logger

logger = get_logger(__name__)
oleCbHdr = 28


class MTEF:
    def __init__(self):
        # mMtefVer     uint8
        self.mMtefVer = 0
        # mPlatform    uint8
        self.mPlatform = 0
        # mProduct     uint8
        self.mProduct = 0
        # mVersion     uint8
        self.mVersion = 0
        # mVersionSub  uint8
        self.mVersionSub = 0
        # mApplication string
        self.mApplication = ''
        # mInline      uint8
        self.mInline = 0

        # reader io.ReadSeeker
        self.reader = None

        # ast   *MtAST
        self.ast = None
        # nodes []*MtAST
        self.nodes = []

        # Valid bool //是否合法，顺利解析
        self.Valid = False

    def readRecord(self):
        """
        读取body的每一行数据并保存到数组里
        """
        # 默认设置为合法的，除非遇到不可解析数据
        self.Valid = True

        # Header
        self.mMtefVer = Helper.bytes2int(self.reader.read(1))  # uint8
        self.mPlatform = Helper.bytes2int(self.reader.read(1))  # uint8
        self.mProduct = Helper.bytes2int(self.reader.read(1))  # uint8
        self.mVersion = Helper.bytes2int(self.reader.read(1))  # uint8
        self.mVersionSub = Helper.bytes2int(self.reader.read(1))  # uint8
        # MTEF v3 没有application key和equation option
        if self.mMtefVer != 3:
            self.mApplication, _ = self.readNullTerminatedString()
            self.mInline = Helper.bytes2int(self.reader.read(1))  # uint8
            self.readBody()
        else:
            self.mApplication = b''
            self.mInline = 0
            self.readBodyV3()

    def readBodyV3(self):
        """
        解析 MTEF v3 主体（单字节 tag：高 4 位＝选项，低 4 位＝类型）
        """
        while True:
            tag_data = self.reader.read(1)
            if not tag_data or len(tag_data) != 1:      # EOF
                break

            tag = Helper.bytes2int(tag_data)
            rec_type = tag & 0x0F                    # 记录类型
            options = (tag & 0xF0) >> 4             # 选项标志

            # END：无附加数据
            if rec_type == RecordTypeV3.END:
                self.nodes.append(MtAST(RecordType.END, None, None))
                continue

            # 需要让下层函数重新读取 tag 的记录
            need_back = rec_type in (
                RecordTypeV3.LINE, RecordTypeV3.CHAR,
                RecordTypeV3.TMPL, RecordTypeV3.PILE,
                RecordTypeV3.MATRIX, RecordTypeV3.EMBELL
            )
            if need_back:
                self.reader.seek(-1, 1)                 # 回退 1 字节

            if rec_type == RecordTypeV3.LINE:
                line = MtLine()
                self.readLineV3(line)
                self.nodes.append(MtAST(RecordTypeV3.LINE, line, None))

            elif rec_type == RecordTypeV3.CHAR:
                ch = MTCharV3()
                self.readCharV3(ch, rec_type)
                self.nodes.append(MtAST(RecordTypeV3.CHAR, ch, None))

            elif rec_type == RecordTypeV3.TMPL:
                tmpl = MtTmpl()
                self.readTMPLV3(tmpl)                     # v3 格式兼容 v5
                self.nodes.append(MtAST(RecordTypeV3.TMPL, tmpl, None))

            elif rec_type == RecordTypeV3.PILE:
                pile = MtPile()
                self.readPileV3(pile)
                self.nodes.append(MtAST(RecordTypeV3.PILE, pile, None))

            elif rec_type == RecordTypeV3.MATRIX:
                mat = MtMatrix()
                self.readMatrixV3(mat)
                self.nodes.append(MtAST(RecordTypeV3.MATRIX, mat, None))

            elif rec_type == RecordTypeV3.EMBELL:
                emb = MtEmbellRd()
                self.readEmbellV3(emb)
                self.nodes.append(MtAST(RecordTypeV3.EMBELL, emb, None))

            elif rec_type == RecordTypeV3.SIZE:         # 2 字节：lsize,dsize
                _ = self.reader.read(2)                 # 跳过即可

            elif rec_type == RecordTypeV3.FULL:
                self.nodes.append(MtAST(RecordTypeV3.FULL, None, None))

            elif rec_type == RecordTypeV3.SUB:
                self.nodes.append(MtAST(RecordTypeV3.SUB, None, None))
            elif rec_type == RecordTypeV3.SUB2:
                self.nodes.append(MtAST(RecordTypeV3.SUB2, None, None))
            elif rec_type == RecordTypeV3.SYM:
                self.nodes.append(MtAST(RecordTypeV3.SYM, None, None))
            elif rec_type == RecordTypeV3.SUBSYM:
                self.nodes.append(MtAST(RecordTypeV3.SUBSYM, None, None))
            else:
                # 未识别记录，标记无效并退出
                self.Valid = False
                break

    def readBody(self):
        while True:
            err = None
            record = RecordType.END
            read_data = None
            read_data = self.reader.read(1)
            if read_data is None or len(read_data) != 1:
                err = 'MEFT.readRecord: read byte error'
            record = Helper.bytes2int(read_data)  # uint8

            # 根据future定义，>=100的后面会跟一个字节，这个字节代表需要跳过的长度
            # For now, readers can assume that an unsigned integer follows the record type and is the number of bytes following it in the record
            # This makes it easy for software that reads MTEF to skip these records.
            if record >= RecordType.FUTURE:
                skipFutureLength = Helper.bytes2int(
                    self.reader.read(1))  # uint8
                self.reader.seek(skipFutureLength, 1)  # io.SeekCurrent
                continue

            if err is not None:
                break

            if record == RecordType.END:
                self.nodes.append(MtAST(RecordType.END, None, None))
            elif record == RecordType.LINE:
                line = MtLine()
                self.readLine(line)

                self.nodes.append(MtAST(RecordType.LINE, line, None))
            elif record == RecordType.CHAR:
                char = MtChar()
                self.readChar(char, record)
                self.nodes.append(MtAST(RecordType.CHAR, char, None))
            elif record == RecordType.TMPL:
                tmpl = MtTmpl()
                self.readTMPL(tmpl)

                self.nodes.append(MtAST(RecordType.TMPL, tmpl, None))
            elif record == RecordType.PILE:
                pile = MtPile()
                self.readPile(pile)

                self.nodes.append(MtAST(RecordType.PILE, pile, None))
            elif record == RecordType.MATRIX:
                matrix = MtMatrix()
                self.readMatrix(matrix)

                self.nodes.append(MtAST(RecordType.MATRIX, matrix, None))

                # 匹配矩阵数据下面的2个nil
                self.nodes.append(MtAST(RecordType.LINE, MtLine(), None))
                self.nodes.append(MtAST(RecordType.LINE, MtLine(), None))
            elif record == RecordType.EMBELL:
                embell = MtEmbellRd()
                self.readEmbell(embell)

                self.nodes.append(MtAST(RecordType.EMBELL, embell, None))
            elif record == RecordType.FONT_STYLE_DEF:
                fsDef = MtfontStyleDef()
                fsDef.fontDefIndex = Helper.bytes2int(
                    self.reader.read(1))  # uint8
                fsDef.name, _ = self.readNullTerminatedString()

                # 读取字节，但是不关心数据，注释
                # m.nodes = append(m.nodes, &MtAST{FONT_STYLE_DEF, fsDef, nil})
            elif record == RecordType.SIZE:
                mtSize = MtSize()
                mtSize.lsize = Helper.bytes2int(self.reader.read(1))  # uint8
                mtSize.dsize = Helper.bytes2int(self.reader.read(1))  # uint8
            elif record == RecordType.SUB:
                self.nodes.append(MtAST(RecordType.SUB, None, None))
            elif record == RecordType.SUB2:
                self.nodes.append(MtAST(RecordType.SUB2, None, None))
            elif record == RecordType.SYM:
                self.nodes.append(MtAST(RecordType.SYM, None, None))
            elif record == RecordType.SUBSYM:
                self.nodes.append(MtAST(RecordType.SUBSYM, None, None))
            elif record == RecordType.FONT_DEF:
                fdef = MtfontDef()
                fdef.encDefIndex = Helper.bytes2int(
                    self.reader.read(1))  # uint8
                fdef.name, _ = self.readNullTerminatedString()

                self.nodes.append(MtAST(RecordType.FONT_DEF, fdef, None))
            elif record == RecordType.COLOR:
                cIndex = MtColorDefIndex()
                cIndex.index = Helper.bytes2int(self.reader.read(1))  # uint8

                # 读取字节，但是不关心数据，注释
                # m.nodes = append(m.nodes, &MtAST{tag: COLOR, value: cIndex, children: nil})
            elif record == RecordType.COLOR_DEF:
                cDef = MtColorDef()
                self.readColorDef(cDef)

                # 读取字节，但是不关心数据，注释
                # m.nodes = append(m.nodes, &MtAST{tag: COLOR_DEF, value: cDef, children: nil})
            elif record == RecordType.FULL:
                self.nodes.append(MtAST(RecordType.FULL, None, None))
            elif record == RecordType.EQN_PREFS:
                prefs = MtEqnPrefs()
                self.readEqnPrefs(prefs)

                self.nodes.append(MtAST(RecordType.EQN_PREFS, prefs, None))
            elif record == RecordType.ENCODING_DEF:
                enc, _ = self.readNullTerminatedString()

                self.nodes.append(MtAST(RecordType.ENCODING_DEF, enc, None))
            else:
                self.Valid = False

        return None

    def readNullTerminatedString(self):
        buf = []
        err = None
        while True:
            p = self.reader.read(1)
            if len(p) != 1:
                err = 'MTEF.readNullTerminatedString.error: read byte error'
            if p[0] == 0:
                break
            buf.append(p)
        return b''.join(buf), err

    def readLine(self, line):
        options = 0  # OptionType
        err = None
        read_data = self.reader.read(1)  # uint8
        if read_data is None or len(read_data) != 1:
            err = 'MTEF.readLine: read byte error'
        options = Helper.bytes2int(read_data)

        if OptionType.MtefOptNudge == OptionType.MtefOptNudge & options:
            line.nudgeX, line.nudgeY, _ = self.readNudge()

        if OptionType.MtefOptLineLspace == OptionType.MtefOptLineLspace & options:
            line.lineSpace = Helper.bytes2int(self.reader.read(1))  # uint8

        # RULER解析
        if OptionType.mtefOPT_LP_RULER == OptionType.mtefOPT_LP_RULER & options:
            # var nStops uint8
            nStops = Helper.bytes2int(self.reader.read(1))  # uint8

            # var tabList []uint8
            tabList = []
            for i in range(nStops):
                stopVal = Helper.bytes2int(self.reader.read(1))  # uint8
                tabList.append(stopVal)

                tabOffset = Helper.bytes2int(self.reader.read(2))  # uint16

        if OptionType.MtefOptLineNull == OptionType.MtefOptLineNull & options:
            line.null = True

        return err

    def readLineV3(self, line):
        """
        读取 MTEF v3 版本的 LINE 记录
        v3 版本的 tag 字节结构：低4位是记录类型，高4位是选项标志
        """
        err = None
        read_data = self.reader.read(1)  # uint8
        if read_data is None or len(read_data) != 1:
            err = 'MTEF.readLineV3: read byte error'
            return err

        tag = Helper.bytes2int(read_data)

        # 从 tag 字节中提取记录类型（低4位）和选项标志（高4位）
        record_type = tag & 0x0F      # 低4位：记录类型
        options = (tag & 0xF0) >> 4   # 高4位右移4位：选项标志

        # 检查各种选项标志（使用按位与运算）

        # 检查 nudge 标志 (xfLMOVE = 0x8)
        if TagTypeV3.xfLMOVE & options:
            line.nudgeX, line.nudgeY, _ = self.readNudgeV3()

        # 检查行间距标志 (xfLSPACE = 0x4)
        if TagTypeV3.xfLSPACE & options:
            line.lineSpace = Helper.bytes2int(self.reader.read(1))  # uint8

        # 检查标尺标志 (xfRULER = 0x2)
        if TagTypeV3.xfRULER & options:
            # 读取 RULER 记录（这是一个完整的记录，不只是标志位后的数据）
            ruler_err = self.readRulerV3(line)
            if ruler_err:
                return ruler_err

        # 检查空行标志 (xfNULL = 0x1)
        if TagTypeV3.xfNULL & options:
            line.null = True

        return err

    def readNudgeV3(self):
        """
        读取 v3 版本的 nudge 值
        根据文档：如果 -128 < dx < +128 且 -128 < dy < +128，
        则存储为两字节，每个值偏移128；否则存储两个128字节，
        后跟16位无偏移值
        """
        first_byte = Helper.bytes2int(self.reader.read(1))
        second_byte = Helper.bytes2int(self.reader.read(1))

        if first_byte == 128 and second_byte == 128:
            # 扩展格式：读取16位值
            nudgeX = Helper.bytes2int(self.reader.read(2))  # int16
            nudgeY = Helper.bytes2int(self.reader.read(2))  # int16
            # 转换为有符号整数（如果需要）
            if nudgeX > 32767:
                nudgeX -= 65536
            if nudgeY > 32767:
                nudgeY -= 65536
        else:
            # 普通格式：值已偏移128
            nudgeX = first_byte - 128
            nudgeY = second_byte - 128

        return nudgeX, nudgeY, None

    def readRulerV3(self, line):
        """
        读取 v3 版本的 RULER 记录
        这是一个完整的记录，需要读取其 tag 字节并处理
        """
        err = None

        # 读取 RULER 记录的 tag 字节
        read_data = self.reader.read(1)
        if read_data is None or len(read_data) != 1:
            return 'MTEF.readRulerV3: read ruler tag error'

        ruler_tag = Helper.bytes2int(read_data)
        ruler_type = ruler_tag & 0x0F      # 应该是 RecordTypeV3.RULER (7)
        ruler_options = (ruler_tag & 0xF0) >> 4

        # 验证记录类型
        if ruler_type != RecordTypeV3.RULER:
            return f'MTEF.readRulerV3: unexpected record type {ruler_type}, expected {RecordTypeV3.RULER}'

        # 检查 RULER 的 nudge（如果有）
        if TagTypeV3.xfLMOVE & ruler_options:
            # RULER 也可能有 nudge 值
            _, _, _ = self.readNudgeV3()  # 暂时忽略 ruler 的 nudge

        # 读取 RULER 的具体数据（根据文档，这部分可能需要进一步实现）
        # 这里暂时跳过，因为文档中没有详细说明 v3 RULER 记录的具体格式

        return err

    def readCharV3(self, char, code):
        """
        读取 MTEF v3 版本的 CHAR 记录
        v3 版本的 tag 字节结构：低4位是记录类型，高4位是选项标志

        与 v5 的主要区别：
        1. v3 使用单个 tag 字节（记录类型+选项标志），v5 使用分离的两个字节
        2. v3 的选项标志值不同：xfAUTO=0x1, xfEMBELL=0x2, xfLMOVE=0x8
        3. v3 的字符编码更简单，通常直接是 ASCII 或扩展字符
        """
        err = None

        # 读取 tag 字节（已在调用方读取并解析）

        # 重新读取 tag 字节来获取选项标志
        # 注意：实际调用时可能需要调整，确保不重复读取
        read_data = self.reader.read(1)
        if read_data is None or len(read_data) != 1:
            return 'MTEF.readCharV3: read tag byte error'

        tag = Helper.bytes2int(read_data)

        # 从 tag 字节中提取选项标志（高4位）
        options = (tag & 0xF0) >> 4   # 高4位右移4位：选项标志

        # 存储选项标志
        char.options = options

        # 检查 nudge 标志 (xfLMOVE = 0x8)
        if TagTypeV3.xfLMOVE & options:
            char.nudgeX, char.nudgeY, _ = self.readNudgeV3()

        # 读取 typeface 值（v3 文档明确说明偏移128）
        typeface_raw = Helper.bytes2int(self.reader.read(1))  # uint8
        if typeface_raw is None:
            return 'MTEF.readCharV3: read typeface error'

        char.typeface = typeface_raw

        # 根据 v3 文档：typeface 偏移 128
        # 正值：MathType 样式 (1=fnTEXT, 2=fnFUNCTION, 3=fnVARIABLE, etc.)
        # 负值：显式字体（由 FONT 记录指定）
        typeface_value = typeface_raw - 128

        # 读取字符代码
        # v3 通常使用 2 字节的字符代码，与 v5 类似
        char.mtcode = Helper.bytes2int(self.reader.read(2))  # uint16
        if char.mtcode is None:
            return 'MTEF.readCharV3: read mtcode error'

        return err

    def readDimensionArrays(self, size):
        shareData = {
            'flag': True,
            'tmpStr': '',
            'count': 0,  # int64
            'array': [],
            'error_count': 0  # 添加错误计数器
        }

        def fx(x):
            # x uint8
            if shareData['flag']:
                if x == 0x00:
                    shareData['flag'] = False
                    shareData['tmpStr'] += 'in'
                elif x == 0x01:
                    shareData['flag'] = False
                    shareData['tmpStr'] += 'cm'
                elif x == 0x02:
                    shareData['flag'] = False
                    shareData['tmpStr'] += 'pt'
                elif x == 0x03:
                    shareData['flag'] = False
                    shareData['tmpStr'] += 'pc'
                elif x == 0x04:
                    shareData['flag'] = False
                    shareData['tmpStr'] += '%'
                else:
                    shareData['error_count'] += 1

            else:
                if x == 0x00:
                    shareData['flag'] = False
                    shareData['tmpStr'] += '0'
                elif x == 0x01:
                    shareData['flag'] = False
                    shareData['tmpStr'] += '1'
                elif x == 0x02:
                    shareData['flag'] = False
                    shareData['tmpStr'] += '2'
                elif x == 0x03:
                    shareData['flag'] = False
                    shareData['tmpStr'] += '3'
                elif x == 0x04:
                    shareData['flag'] = False
                    shareData['tmpStr'] += '4'
                elif x == 0x05:
                    shareData['flag'] = False
                    shareData['tmpStr'] += '5'
                elif x == 0x06:
                    shareData['flag'] = False
                    shareData['tmpStr'] += '6'
                elif x == 0x07:
                    shareData['flag'] = False
                    shareData['tmpStr'] += '7'
                elif x == 0x08:
                    shareData['flag'] = False
                    shareData['tmpStr'] += '8'
                elif x == 0x09:
                    shareData['flag'] = False
                    shareData['tmpStr'] += '9'
                elif x == 0x0a:
                    shareData['flag'] = False
                    shareData['tmpStr'] += '.'
                elif x == 0x0b:
                    shareData['flag'] = False
                    shareData['tmpStr'] += '-'
                elif x == 0x0f:
                    shareData['flag'] = True
                    shareData['count'] += 1
                    shareData['array'].append(shareData['tmpStr'])
                    shareData['tmpStr'] = ''
                else:
                    shareData['error_count'] += 1

        max_iterations = size * 10  # 设置最大迭代次数保护
        iteration_count = 0

        while True:
            if shareData['count'] >= size:
                break

            # 防止无限循环
            iteration_count += 1
            if iteration_count > max_iterations:

                break

            # 如果错误过多，提前退出
            if shareData['error_count'] > 50:
                break

            # 检查是否还有数据可读
            try:
                ch_data = self.reader.read(1)
                if not ch_data or len(ch_data) == 0:
                    break
                ch = Helper.bytes2int(ch_data)
            except Exception as e:
                break

            hi = int((ch & 0xf0) / 16)
            lo = ch & 0x0f
            fx(hi)
            fx(lo)

        return shareData['array'], None

    def readEqnPrefs(self, eqnPrefs):
        options = Helper.bytes2int(self.reader.read(1))

        # sizes
        size = Helper.bytes2int(self.reader.read(1))
        eqnPrefs.sizes, _ = self.readDimensionArrays(size)

        # spaces
        size = 0
        size = Helper.bytes2int(self.reader.read(1))
        eqnPrefs.spaces, _ = self.readDimensionArrays(size)

        # styles
        size = 0
        size = Helper.bytes2int(self.reader.read(1))
        styles = []  # byte
        for _ in range(size):
            c = 0  # uint8
            c = Helper.bytes2int(self.reader.read(1))
            if c == 0:
                styles.append(0)
            else:
                c = Helper.bytes2int(self.reader.read(1))
                styles.append(c)
        eqnPrefs.styles = styles
        return None

    def readChar(self, char, record):
        options = Helper.bytes2int(self.reader.read(1))  # uint8

        # ---------- v5/v4 解析 ----------
        if OptionType.MtefOptNudge == OptionType.MtefOptNudge & options:
            char.nudgeX, char.nudgeY, _ = self.readNudge()

        char.typeface = Helper.bytes2int(self.reader.read(1))  # uint8

        if OptionType.MtefOptCharEncNoMtcode != OptionType.MtefOptCharEncNoMtcode & options:
            char.mtcode = Helper.bytes2int(self.reader.read(2))  # uint16

        if OptionType.MtefOptCharEncChar8 == OptionType.MtefOptCharEncChar8 & options:
            char.bits8 = Helper.bytes2int(self.reader.read(1))  # uint8
        if OptionType.MtefOptCharEncChar16 == OptionType.MtefOptCharEncChar16 & options:
            char.bits16 = Helper.bytes2int(self.reader.read(2))  # uint16

        return None

    def readNudge(self):
        b1 = Helper.bytes2int(self.reader.read(2))  # 类型有待确认
        b2 = Helper.bytes2int(self.reader.read(2))

        err = None
        if b1 == 128 or b2 == 128:
            nudgeX = Helper.bytes2int(self.reader.read(2))  # int16
            nudgeY = Helper.bytes2int(self.reader.read(2))  # int16
            return nudgeX, nudgeY, err
        else:
            nudgeX = b1
            nudgeY = b2
            return nudgeX, nudgeY, err

    def readTMPLV3(self, tmpl):
        """
        读取 MTEF v3 版本的 TMPL 记录
        v3 和 v5 的主要差异：
        1. v3 没有单独的 options 字节，选项标志在 tag 字节的高4位
        2. v3 的结构更简单，没有一些 v5 新增的特性
        """
        err = None

        # 读取 tag 字节
        read_data = self.reader.read(1)
        if read_data is None or len(read_data) != 1:
            return 'MTEF.readTMPLV3: read tag error'

        tag = Helper.bytes2int(read_data)
        record_type = tag & 0x0F      # 低4位：记录类型
        options = (tag & 0xF0) >> 4   # 高4位：选项标志

        # 验证记录类型
        if record_type != RecordTypeV3.TMPL:
            return f'MTEF.readTMPLV3: unexpected record type {record_type}, expected {RecordTypeV3.TMPL}'

        # 检查 nudge 标志 (xfLMOVE = 0x8)
        if TagTypeV3.xfLMOVE & options:
            tmpl.nudgeX, tmpl.nudgeY, _ = self.readNudgeV3()

        # 读取 selector（模板选择器代码）
        tmpl.selector = Helper.bytes2int(self.reader.read(1))  # uint8

        # 读取 variation（模板变体代码）
        # v3 版本的 variation 读取方式与 v5 类似
        byte1 = Helper.bytes2int(self.reader.read(1))
        if 0x80 == byte1 & 0x80:
            # 如果第一个字节的最高位设置，则读取第二个字节
            byte2 = Helper.bytes2int(self.reader.read(1))
            tmpl.variation = (byte1 & 0x7F) | (byte2 << 8)
        else:
            tmpl.variation = byte1

        # 读取 options（模板特定选项）
        # v3 版本也有这个字段，主要用于积分和围栏模板
        tmpl.options = Helper.bytes2int(self.reader.read(1))  # uint8

        return err

    def readTMPL(self, tmpl):
        options = 0  # OptionType
        options = Helper.bytes2int(self.reader.read(1))

        if OptionType.MtefOptNudge == OptionType.MtefOptNudge & options:
            tmpl.nudgeX, tmpl.nudgeY, _ = self.readNudge()

        tmpl.selector = Helper.bytes2int(self.reader.read(1))  # uint8

        # variation, 1 or 2 bytes
        byte1 = 0  # uint8
        byte1 = Helper.bytes2int(self.reader.read(1))
        if 0x80 == byte1 & 0x80:
            byte2 = 0  # uint8
            byte2 = Helper.bytes2int(self.reader.read(1))
            # tmpl.variation = (uint16(byte1) & 0x7F) | (uint16(byte2) << 8)
            tmpl.variation = (byte1 & 0x7F) | (byte2 << 8)
        else:
            tmpl.variation = byte1
        tmpl.options = Helper.bytes2int(self.reader.read(1))  # uint8
        return None

    def readPileV3(self, pile):
        """
        读取 MTEF v3 版本的 PILE 记录
        v3 版本的 PILE 记录结构：
        - tag (4) [nudge if xfLMOVE] [halign] [valign] [RULER if xfRULER] [object list]
        """
        err = None

        # 读取 tag 字节
        read_data = self.reader.read(1)
        if read_data is None or len(read_data) != 1:
            return 'MTEF.readPileV3: read tag error'

        tag = Helper.bytes2int(read_data)
        record_type = tag & 0x0F      # 低4位：记录类型
        options = (tag & 0xF0) >> 4   # 高4位：选项标志

        # 验证记录类型
        if record_type != RecordTypeV3.PILE:
            return f'MTEF.readPileV3: unexpected record type {record_type}, expected {RecordTypeV3.PILE}'

        # 检查 nudge 标志 (xfLMOVE = 0x8)
        if TagTypeV3.xfLMOVE & options:
            pile.nudgeX, pile.nudgeY, _ = self.readNudgeV3()

        # 读取水平对齐方式 (halign)
        # 1=左对齐, 2=居中, 3=右对齐, 4=关系运算符对齐, 5=小数点对齐
        pile.halign = Helper.bytes2int(self.reader.read(1))  # uint8

        # 读取垂直对齐方式 (valign)
        # 0=与顶行基线对齐, 1=与中心行基线对齐, 2=与底行基线对齐, 3=垂直居中
        pile.valign = Helper.bytes2int(self.reader.read(1))  # uint8

        # 检查标尺标志 (xfRULER = 0x2)
        if TagTypeV3.xfRULER & options:
            # 读取 RULER 记录（这是一个完整的记录）
            ruler_err = self.readRulerV3(pile)
            if ruler_err:
                return ruler_err

        return err

    def readPile(self, pile):
        options = 0  # OptionType
        options = Helper.bytes2int(self.reader.read(1))

        if OptionType.MtefOptNudge == OptionType.MtefOptNudge & options:
            pile.nudgeX, pile.nudgeY, _ = self.readNudge()

        # 读取halign和valign
        pile.halign = Helper.bytes2int(self.reader.read(1))  # uint8
        pile.valign = Helper.bytes2int(self.reader.read(1))  # uint8

        return None

    def readMatrixV3(self, matrix):
        """
        读取 MTEF v3 版本的 MATRIX 记录
        v3 版本的 MATRIX 记录结构：
        - tag (5) [nudge if xfLMOVE] [valign] [h_just] [v_just] [rows] [cols]
          [row_parts] [col_parts] [object list]
        """
        err = None

        # 读取 tag 字节
        read_data = self.reader.read(1)
        if read_data is None or len(read_data) != 1:
            return 'MTEF.readMatrixV3: read tag error'

        tag = Helper.bytes2int(read_data)
        record_type = tag & 0x0F      # 低4位：记录类型
        options = (tag & 0xF0) >> 4   # 高4位：选项标志

        # 验证记录类型
        if record_type != RecordTypeV3.MATRIX:
            return f'MTEF.readMatrixV3: unexpected record type {record_type}, expected {RecordTypeV3.MATRIX}'

        # 检查 nudge 标志 (xfLMOVE = 0x8)
        if TagTypeV3.xfLMOVE & options:
            matrix.nudgeX, matrix.nudgeY, _ = self.readNudgeV3()

        # 读取矩阵在容器内的垂直对齐方式
        matrix.valign = Helper.bytes2int(self.reader.read(1))  # uint8

        # 读取列内的水平对齐方式
        matrix.h_just = Helper.bytes2int(self.reader.read(1))  # uint8

        # 读取列内的垂直对齐方式
        matrix.v_just = Helper.bytes2int(self.reader.read(1))  # uint8

        # 读取行数和列数
        matrix.rows = Helper.bytes2int(self.reader.read(1))  # uint8
        matrix.cols = Helper.bytes2int(self.reader.read(1))  # uint8

        # 读取行分隔线类型列表
        # 每个可能的分隔线（比行数多1）占用2位，舍入到最近的字节
        # 每个值确定相应分隔线的样式（0=无, 1=实线, 2=虚线, 3=点线）
        row_parts_bytes = (matrix.rows + 1 + 3) // 4  # 每字节4个2位值，向上取整
        for _ in range(row_parts_bytes):
            Helper.bytes2int(self.reader.read(1))  # 暂时读取但不存储

        # 读取列分隔线类型列表（类似行分隔线）
        col_parts_bytes = (matrix.cols + 1 + 3) // 4  # 每字节4个2位值，向上取整
        for _ in range(col_parts_bytes):
            Helper.bytes2int(self.reader.read(1))  # 暂时读取但不存储

        return err

    def readMatrix(self, matrix):
        options = 0  # OptionType
        options = Helper.bytes2int(self.reader.read(1))

        if OptionType.MtefOptNudge == OptionType.MtefOptNudge & options:
            matrix.nudgeX, matrix.nudgeY, _ = self.readNudge()

        # 读取valign和h_just、v_just
        matrix.valign = Helper.bytes2int(self.reader.read(1))  # uint8
        matrix.h_just = Helper.bytes2int(self.reader.read(1))  # uint8
        matrix.v_just = Helper.bytes2int(self.reader.read(1))  # uint8

        # 读取rows和cols
        matrix.rows = Helper.bytes2int(self.reader.read(1))  # uint8
        matrix.cols = Helper.bytes2int(self.reader.read(1))  # uint8

        return None

    def readEmbell(self, embell):
        options = 0  # OptionType
        options = Helper.bytes2int(self.reader.read(1))

        if OptionType.MtefOptNudge == OptionType.MtefOptNudge & options:
            embell.nudgeX, embell.nudgeY, _ = self.readNudge()

        # 读取embellishment type
        embell.embellType = Helper.bytes2int(self.reader.read(1))  # uint8

        return None

    def readEmbellV3(self, embell):
        """
        读取 MTEF v3 版本的 EMBELL 记录
        v3 版本的 EMBELL 记录结构：
        - tag (6) [nudge if xfLMOVE] [embell]

        v3 版本的装饰类型与 v5 稍有不同，使用不同的常量
        """
        err = None

        # 读取 tag 字节
        read_data = self.reader.read(1)
        if read_data is None or len(read_data) != 1:
            return 'MTEF.readEmbellV3: read tag error'

        tag = Helper.bytes2int(read_data)
        record_type = tag & 0x0F      # 低4位：记录类型
        options = (tag & 0xF0) >> 4   # 高4位：选项标志

        # 验证记录类型
        if record_type != RecordTypeV3.EMBELL:
            return f'MTEF.readEmbellV3: unexpected record type {record_type}, expected {RecordTypeV3.EMBELL}'

        # 检查 nudge 标志 (xfLMOVE = 0x8)
        if TagTypeV3.xfLMOVE & options:
            embell.nudgeX, embell.nudgeY, _ = self.readNudgeV3()

        # 读取装饰类型
        # v3 版本的装饰类型值（参考官网文档中的 EmbellTypeV3）：
        # 2=单点, 3=双点, 4=三点, 5=单撇, 6=双撇, 7=反向撇, 8=波浪线,
        # 9=帽子, 10=斜杠, 11=右箭头, 12=左箭头, 13=双向箭头, 等等
        embell.embellType = Helper.bytes2int(self.reader.read(1))  # uint8

        return err

    def readColorDef(self, colorDef):
        options = 0  # OptionType
        options = Helper.bytes2int(self.reader.read(1))

        color = 0  # uint16
        if OptionType.mtefCOLOR_CMYK == OptionType.mtefCOLOR_CMYK & options:
            # CMYK，读4个值
            for _ in range(4):
                color = Helper.bytes2int(self.reader.read(2))
                colorDef.values.append(color)
        else:
            # RGB，读3个值
            for _ in range(3):
                color = Helper.bytes2int(self.reader.read(2))
                colorDef.values.append(color)

        if OptionType.mtefCOLOR_NAME == OptionType.mtefCOLOR_NAME & options:
            colorDef.name, _ = self.readNullTerminatedString()

        return None

    def Translate(self):
        if self.mMtefVer != 3:
            latexStr, err = self.makeLatex(self.ast)

        else:
            # v3没有root节点，所以需要手动添加
            latexStr, err = self.makeLatexV3(self.ast)
            latexStr = self.fixConsecutiveScripts(latexStr)
            format = ["$", latexStr, "$"]
            latexStr = "".join(format)

        if err is not None:
            logger.error('MTEF.Translate.err: %s', err)

        if self.Valid:
            return latexStr
        else:
            return ''

    def fixConsecutiveScripts(self, latex_str):
        """
        修复连续的下标和上标问题
        将形如 X_A_B 或 X^A^B 的连续脚本转换为 X_{AB} 或 X^{AB}
        """
        import re

        # 修复连续下标：合并相邻的下标
        # 先处理带花括号的下标：_{ content1 }_{ content2 } -> _{content1 content2}
        pattern1 = r'_\{\s*([^}]+)\s*\}\s*_\{\s*([^}]+)\s*\}'
        while re.search(pattern1, latex_str):
            latex_str = re.sub(pattern1, r'_{\1 \2}', latex_str)

        # 再处理简单下标：_A_B -> _{AB}
        pattern2 = r'_([^_^{}\s\\])\s*_([^_^{}\s\\])'
        while re.search(pattern2, latex_str):
            latex_str = re.sub(pattern2, r'_{\1\2}', latex_str)

        # 修复连续上标：合并相邻的上标
        # 先处理带花括号的上标：^{ content1 }^{ content2 } -> ^{content1 content2}
        pattern3 = r'\^\{\s*([^}]+)\s*\}\s*\^\{\s*([^}]+)\s*\}'
        while re.search(pattern3, latex_str):
            latex_str = re.sub(pattern3, r'^{\1 \2}', latex_str)

        # 再处理简单上标：^A^B -> ^{AB}
        pattern4 = r'\^([^_^{}\s\\])\s*\^([^_^{}\s\\])'
        while re.search(pattern4, latex_str):
            latex_str = re.sub(pattern4, r'^{\1\2}', latex_str)

        return latex_str

    def makeAST(self):
        """
        根据数组生成出栈入栈结构
        """
        ast = MtAST()
        ast.tag = 0xff
        ast.value = None
        self.ast = ast

        # v3 格式需要特殊处理
        if self.mMtefVer == 3:
            return self.makeASTv3()

        stack = []
        stack.append(ast)

        for node in self.nodes:
            if node.tag == RecordType.LINE:
                if len(stack):
                    parent = stack[len(stack) - 1]
                    if not parent.children:
                        parent.children = []
                    parent.children.append(node)
                if not node.value.null:
                    stack.append(node)
            if node.tag == RecordType.TMPL:
                if len(stack):
                    parent = stack[len(stack) - 1]
                    if not parent.children:
                        parent.children = []
                    parent.children.append(node)

                stack.append(node)
            if node.tag == RecordType.PILE:
                if len(stack):
                    parent = stack[len(stack) - 1]
                    if not parent.children:
                        parent.children = []
                    parent.children.append(node)

                stack.append(node)
            if node.tag == RecordType.MATRIX:
                if len(stack):
                    parent = stack[len(stack) - 1]
                    if not parent.children:
                        parent.children = []
                    parent.children.append(node)

                stack.append(node)
            if node.tag == RecordType.END:
                if len(stack):
                    ele = stack[len(stack) - 1]
                    stack.remove(ele)
            if node.tag == RecordType.CHAR:
                if len(stack):
                    parent = stack[len(stack) - 1]
                    if not parent.children:
                        parent.children = []
                    parent.children.append(node)
                else:
                    ast.children.append(node)
            if node.tag == RecordType.EMBELL:
                if len(stack):
                    parent = stack[len(stack) - 1]
                    if not parent.children:
                        parent.children = []
                    parent.children.append(node)

                    embellType = node.value.embellType
                    if embellType == EmbellType.emb1DOT or embellType == EmbellType.embHAT or embellType == EmbellType.embOBAR:
                        if len(parent.children) >= 2:
                            embellData = parent.children[len(parent.children) - 1]
                            charData = parent.children[len(parent.children) - 2]
                            parent.children = parent.children[:len(parent.children) - 2]

                            parent.children.append(embellData)
                            parent.children.append(charData)

                stack.append(node)

        return None

    def makeASTv3(self):
        """
        专门处理 v3 格式的 AST 构建
        """
        stack = []
        stack.append(self.ast)

        for node in self.nodes:
            if node.tag == RecordTypeV3.LINE:
                if len(stack):
                    parent = stack[len(stack) - 1]
                    if not parent.children:
                        parent.children = []
                    parent.children.append(node)
                if not node.value.null:
                    stack.append(node)
            if node.tag == RecordTypeV3.TMPL:
                if len(stack):
                    parent = stack[len(stack) - 1]
                    if not parent.children:
                        parent.children = []
                    parent.children.append(node)

                stack.append(node)
            if node.tag == RecordTypeV3.PILE:
                if len(stack):
                    parent = stack[len(stack) - 1]
                    if not parent.children:
                        parent.children = []
                    parent.children.append(node)

                stack.append(node)
            if node.tag == RecordTypeV3.MATRIX:
                if len(stack):
                    parent = stack[len(stack) - 1]
                    if not parent.children:
                        parent.children = []
                    parent.children.append(node)

                stack.append(node)
            if node.tag == RecordTypeV3.END:
                if len(stack):
                    ele = stack[len(stack) - 1]
                    stack.remove(ele)
            if node.tag == RecordTypeV3.CHAR:
                if len(stack):
                    parent = stack[len(stack) - 1]
                    if not parent.children:
                        parent.children = []
                    parent.children.append(node)
                else:
                    self.ast.children.append(node)
            if node.tag == RecordTypeV3.EMBELL:
                if len(stack):
                    parent = stack[len(stack) - 1]
                    if not parent.children:
                        parent.children = []
                    parent.children.append(node)

                    embellType = node.value.embellType
                    if (embellType == EmbellTypeV3.embDOT or embellType == EmbellTypeV3.embHAT
                            or embellType == EmbellTypeV3.embOBAR):
                        if len(parent.children) >= 2:
                            embellData = parent.children[len(
                                parent.children) - 1]
                            charData = parent.children[len(
                                parent.children) - 2]
                            parent.children = parent.children[:len(
                                parent.children) - 2]

                            parent.children.append(embellData)
                            parent.children.append(charData)

                stack.append(node)

        return None

    def makeLatex(self, ast):
        """
        根据出栈入栈结构生成latex字符串
        """

        buf = ''

        if ast.tag == RecordType.ROOT:
            buf += '$ '
            for _ast in ast.children:
                _latex, _ = self.makeLatex(_ast)
                buf += _latex
            buf += ' $'
            return buf, None
        elif ast.tag == RecordType.CHAR:
            mtcode = ast.value.mtcode
            typeface = ast.value.typeface
            char = chr(mtcode)

            # 生成char的一些特殊集
            hexExtend = ''
            typefaceFmt = ''
            if typeface - 128 == CharTypeface.fnMTEXTRA:
                hexExtend = '/mathmode'
            elif typeface - 128 == CharTypeface.fnSPACE:
                hexExtend = '/mathmode'
            elif typeface - 128 == CharTypeface.fnTEXT:
                typefaceFmt = '{ \\rm{ %s } }'

            # 生成扩展字符的key
            # hexCode := fmt.Sprintf("%04x", mtcode)
            hexCode = '%04x' % mtcode
            # hexKey := fmt.Sprintf("char/0x%s%s", hexCode, hexExtend)
            hexKey = 'char/0x%s%s' % (hexCode, hexExtend)

            # fmt.Println(char, hexKey)

            # 首先去找扩展字符
            sChar = Chars.get(hexKey)
            if sChar:
                char = sChar
            else:
                # 如果char是特殊symbol，需要转义
                sChar = SpecialChar.get(char)
                if sChar:
                    char = sChar

            # 确定字符是否为文本，如果是文本，则需要包一层
            if typefaceFmt != '':
                char = typefaceFmt % char

            buf += char
            return buf, None
        elif ast.tag == RecordType.TMPL:
            # 强制类型转换为MtTmpl
            tmpl = ast.value

            if tmpl.selector == SelectorType.tmANGLE:
                if len(ast.children) < 3:
                    return '<unknown>', None
                mainAST = ast.children[0]
                leftAST = ast.children[1]
                rightAST = ast.children[2]

                mainSlot, _ = self.makeLatex(mainAST)
                leftSlot, _ = self.makeLatex(leftAST)
                rightSlot, _ = self.makeLatex(rightAST)

                # 转成latex代码
                mainStr = ''
                leftStr = ''
                rightStr = ''
                if mainSlot != '':
                    mainStr = '{ %s }' % mainSlot
                if leftSlot != '':
                    leftStr = '\\left %s' % leftSlot
                if rightSlot != '':
                    rightStr = '\\right %s' % rightSlot

                buf += '%s %s %s' % (leftStr, mainStr, rightStr)

                return buf, None
            elif tmpl.selector == SelectorType.tmPAREN:
                if len(ast.children) < 3:
                    return '()', None
                mainAST = ast.children[0]
                leftAST = ast.children[1]
                rightAST = ast.children[2]

                mainSlot, _ = self.makeLatex(mainAST)
                leftSlot, _ = self.makeLatex(leftAST)
                rightSlot, _ = self.makeLatex(rightAST)

                # 转成latex代码
                mainStr = ''
                leftStr = ''
                rightStr = ''
                if mainSlot != '':
                    mainStr = '{ %s }' % mainSlot
                if leftSlot != '':
                    leftStr = '\\left %s' % leftSlot
                if rightSlot != '':
                    rightStr = '\\right %s' % rightSlot

                buf += '%s %s %s' % (leftStr, mainStr, rightStr)
                return buf, None
            elif tmpl.selector == SelectorType.tmBRACE:
                mainSlot = ''
                leftSlot = ''
                rightSlot = ''
                idx = 0
                for astData in ast.children:
                    if idx == 0:
                        mainSlot, _ = self.makeLatex(astData)
                    elif idx == 1:
                        leftSlot, _ = self.makeLatex(astData)
                    else:
                        rightSlot, _ = self.makeLatex(astData)
                    idx += 1

                if rightSlot == '':
                    rightSlot = '.'
                else:
                    rightSlot = ' ' + rightSlot

                # 组装公式
                buf += '\\left %s \\begin{array}{l} %s \\end{array} \\right%s' % (leftSlot, mainSlot, rightSlot)

                return buf, None
            elif tmpl.selector == SelectorType.tmBRACK:
                mainAST = ast.children[0]
                leftAST = ast.children[1]
                rightAST = ast.children[2]
                mainSlot, _ = self.makeLatex(mainAST)
                if mainSlot == '':
                    mainSlot = '\\space'
                leftSlot, _ = self.makeLatex(leftAST)
                rightSlot, _ = self.makeLatex(rightAST)
                buf += '\\left%s %s \\right%s' % (leftSlot, mainSlot, rightSlot)
                return buf, None
            elif tmpl.selector == SelectorType.tmBAR:
                # 读取数据 ParBoxClass
                mainSlot = ''
                leftSlot = ''
                rightSlot = ''
                idx = 0
                for astData in ast.children:
                    if idx == 0:
                        mainSlot, _ = self.makeLatex(astData)
                    elif idx == 1:
                        leftSlot, _ = self.makeLatex(astData)
                    else:
                        rightSlot, _ = self.makeLatex(astData)
                    idx += 1

                if rightSlot == '':
                    rightSlot = '.'
                else:
                    rightSlot = ' ' + rightSlot

                # 转成latex代码
                mainStr = ''
                leftStr = ''
                rightStr = ''
                if mainSlot != '':
                    mainStr = '{ %s }' % mainSlot
                if leftSlot != '':
                    leftStr = '\\left %s' % leftSlot
                if rightSlot != '':
                    rightStr = '\\right %s' % rightSlot

                # 组成整体公式
                tmplStr = '%s %s %s' % (leftStr, mainStr, rightStr)
                buf += tmplStr

                return buf, None
            elif tmpl.selector == SelectorType.tmINTERVAL:
                # 读取数据 ParBoxClass
                mainAST = ast.children[0]
                leftAST = ast.children[1]
                rightAST = ast.children[2]

                # 读取latex数据
                mainSlot, _ = self.makeLatex(mainAST)
                leftSlot, _ = self.makeLatex(leftAST)
                rightSlot, _ = self.makeLatex(rightAST)

                # 转成latex代码
                mainStr = ''
                leftStr = ''
                rightStr = ''
                if mainSlot != '':
                    mainStr = '{ %s }' % mainSlot
                if leftSlot != '':
                    leftStr = '\\left %s' % leftSlot
                if rightSlot != '':
                    rightStr = '\\right %s' % rightSlot

                # 组成整体公式
                tmplStr = '%s %s %s' % (leftStr, mainStr, rightStr)
                buf += tmplStr

                return buf, None
            elif tmpl.selector == SelectorType.tmROOT:
                mainAST = ast.children[0]
                radiAST = ast.children[1]
                mainSlot, _ = self.makeLatex(mainAST)
                radiSlot, _ = self.makeLatex(radiAST)
                buf += '\\sqrt[%s] { %s }' % (radiSlot, mainSlot)
                return buf, None
            elif tmpl.selector == SelectorType.tmFRACT:
                if len(ast.children) < 2:
                    # 直接传入bin文件这里不会触发，传入字节流在公式太多/thesis_06公式太多.docx以及thesis_15公式太多中会触发
                    numAST = ast.children[0]
                    numSlot, _ = self.makeLatex(numAST)
                    buf += '\\frac { %s } {Unknown}' % numSlot
                    return buf, None
                numAST = ast.children[0]
                denAST = ast.children[1]
                numSlot, _ = self.makeLatex(numAST)
                denSlot, _ = self.makeLatex(denAST)
                buf += '\\frac { %s } { %s }' % (numSlot, denSlot)
                return buf, None
            elif tmpl.selector == SelectorType.tmARROW:
                """
                    variation	symbol	description
                    0×0000	tvAR_SINGLE	single arrow
                    0×0001	tvAR_DOUBLE	double arrow
                    0×0002	tvAR_HARPOON	harpoon
                    0×0004	tvAR_TOP	top slot is present
                    0×0008	tvAR_BOTTOM	bottom slot is present
                    0×0010	tvAR_LEFT	if single, arrow points left
                    0×0020	tvAR_RIGHT	if single, arrow points right
                    0×0010	tvAR_LOS	if double or harpoon, large over small
                    0×0020	tvAR_SOL	if double or harpoon, small over large
                """
                topAST = ast.children[0]
                bottomAST = ast.children[1]

                # 读取latex数据
                topSlot, _ = self.makeLatex(topAST)
                bottomSlot, _ = self.makeLatex(bottomAST)

                # 转成latex代码
                topStr = ''
                bottomStr = ''
                if topSlot != '':
                    topStr = '{\\mathrm{ %s }}' % topSlot
                if bottomSlot != '':
                    bottomStr = '[\\mathrm{ %s }]' % bottomSlot

                """
                    variation转码
                """
                variationsMap = {}  # map[uint16]string
                variationsMap[0x0000] = "single"
                variationsMap[0x0001] = "double"
                variationsMap[0x0002] = "harpoon"
                variationsMap[0x0004] = "topSlotPresent"
                variationsMap[0x0008] = "bottomSlotPresent"
                variationsMap[0x0010] = "pointLeft"
                variationsMap[0x0020] = "pointRight"

                # 有序循环
                variationsCode = [0x0000, 0x0001, 0x0002, 0x0004, 0x0008, 0x0010,
                                  0x0020]  # []uint16{0x0000, 0x0001, 0x0002, 0x0004, 0x0008, 0x0010, 0x0020}

                arrowStyle = "single"
                latexFmt = "\\x"
                for vCode in variationsCode:
                    # 如果存在掩码
                    if vCode & tmpl.variation != 0:
                        # 判断类型，默认是single
                        if variationsMap[vCode] == "double":
                            arrowStyle = "double"
                        elif variationsMap[vCode] == "harpoon":
                            arrowStyle = "harpoon"

                        if arrowStyle == "single" and variationsMap[vCode] == "pointLeft":
                            latexFmt = latexFmt + "leftarrow"
                        elif arrowStyle == "double" and variationsMap[vCode] == "pointLeft":
                            logger.warning('MTEF.makeLatex: not implement double , large over small')
                        elif arrowStyle == "harpoon" and variationsMap[vCode] == "pointLeft":
                            logger.warning('MTEF.makeLatex: not implement harpoon, large over small')

                        if arrowStyle == "single" and variationsMap[vCode] == "pointRight":
                            latexFmt = latexFmt + "rightarrow"
                        elif arrowStyle == "double" and variationsMap[vCode] == "pointRight":
                            logger.warning('MTEF.makeLatex: not implement double , small over large')
                        elif arrowStyle == "harpoon" and variationsMap[vCode] == "pointRight":
                            logger.warning('MTEF.makeLatex: not implement harpoon, small over large')

                # 组成整体公式
                tmplStr = "%s %s %s" % (latexFmt, bottomStr, topStr)
                buf += tmplStr

                return buf, None
            elif tmpl.selector == SelectorType.tmUBAR:
                # 读取数据
                mainAST = ast.children[0]

                # 读取latex数据
                mainSlot, _ = self.makeLatex(mainAST)

                # 转成latex代码
                mainStr = ''
                if mainSlot != "":
                    mainStr = " {\\underline{ %s }} " % mainSlot

                # 组成整体公式
                tmplStr = " %s " % mainStr
                buf += tmplStr

                # 返回数据
                return buf, None
            elif tmpl.selector == SelectorType.tmOBAR:
                # 读取数据 - 上划线 (overbar)
                mainAST = ast.children[0]

                # 读取latex数据
                mainSlot, _ = self.makeLatex(mainAST)

                # 转成latex代码
                mainStr = ''
                if mainSlot != "":
                    # 检查variation来决定是否使用双线
                    if tmpl.variation & 0x0001:  # tvBAR_DOUBLE - 双划线
                        mainStr = " {\\overline{\\overline{ %s }}} " % mainSlot
                    else:  # 单划线
                        mainStr = " {\\overline{ %s }} " % mainSlot

                # 组成整体公式
                tmplStr = " %s " % mainStr
                buf += tmplStr

                # 返回数据
                return buf, None
            elif tmpl.selector == SelectorType.tmSUM:
                # 读取数据 BigOpBoxClass
                mainSlot = ''
                upperSlot = ''
                lowerSlot = ''
                operatorSlot = ''
                idx = 0
                for astData in ast.children:
                    if idx == 0:
                        mainSlot, _ = self.makeLatex(astData)
                    elif idx == 1:
                        lowerSlot, _ = self.makeLatex(astData)
                    elif idx == 2:
                        upperSlot, _ = self.makeLatex(astData)
                    else:
                        operatorSlot, _ = self.makeLatex(astData)
                    idx += 1

                # 转成latex代码
                mainStr = ''
                lowerStr = ''
                upperStr = ''
                if mainSlot != "":
                    mainStr = "{ %s }" % mainSlot
                if lowerSlot != "":
                    lowerStr = "\\limits_{ %s }" % lowerSlot
                if upperSlot != "":
                    upperStr = "^ %s" % upperSlot

                # 组成整体公式
                tmplStr = "%s %s %s %s" % (operatorSlot, lowerStr, upperStr, mainStr)
                buf += tmplStr

                return buf, None
            elif tmpl.selector == SelectorType.tmLIM:
                # 读取数据 LimBoxClass
                mainSlot = ''
                lowerSlot = ''
                upperSlot = ''
                idx = 0
                for astData in ast.children:
                    if idx == 0:
                        mainSlot, _ = self.makeLatex(astData)
                    elif idx == 1:
                        lowerSlot, _ = self.makeLatex(astData)
                    else:
                        upperSlot, _ = self.makeLatex(astData)
                    idx += 1

                # 转成latex代码
                mainStr = ''
                lowerStr = ''
                upperStr = ''
                if mainSlot != "":
                    mainStr = "\\mathop { %s }" % mainSlot
                if lowerSlot != "":
                    lowerStr = "\\limits_{ %s }" % lowerSlot
                if upperSlot != "":
                    upperStr = ""

                # 组成整体公式
                tmplStr = "%s %s %s" % (mainStr, lowerStr, upperStr)
                buf += tmplStr

                return buf, None
            elif tmpl.selector == SelectorType.tmSUP:
                # 只处理上标 (superscript only)
                if ast.children:
                    supAST = ast.children[0]  # 只读取第一个子对象作为上标内容
                    supSlot, _ = self.makeLatex(supAST)

                    if supSlot:
                        buf += f"^{{ {supSlot} }}"
                return buf, None
            elif tmpl.selector == SelectorType.tmSUB:
                # 只处理下标 (subscript only)
                if ast.children:
                    subAST = ast.children[0]  # 只读取第一个子对象作为下标内容
                    subSlot, _ = self.makeLatex(subAST)

                    if subSlot:
                        buf += f"_{{ {subSlot} }}"
                return buf, None
            elif tmpl.selector == SelectorType.tmSUBSUP:
                # 同时处理下标和上标 (both subscript and superscript)
                subSlot = ''
                supSlot = ''

                if len(ast.children) >= 1:
                    subAST = ast.children[0]  # 第一个子对象是下标
                    subSlot, _ = self.makeLatex(subAST)

                if len(ast.children) >= 2:
                    supAST = ast.children[1]  # 第二个子对象是上标
                    supSlot, _ = self.makeLatex(supAST)

                # 转成latex代码
                subFmt = f"_{{ {subSlot} }}" if subSlot else ''
                supFmt = f"^{{ {supSlot} }}" if supSlot else ''

                # 组成整体公式
                tmplStr = f"{subFmt}{supFmt}"
                buf += tmplStr
                return buf, None
            elif tmpl.selector == SelectorType.tmVEC:
                """
                    variations：
                    variation	symbol	description
                    0×0001	tvVE_LEFT	arrow points left
                    0×0002	tvVE_RIGHT	arrow points right
                    0×0004	tvVE_UNDER	arrow under slot, else over slot
                    0×0008	tvVE_HARPOON	harpoon

                    这个转换是通过掩码计算的：
                    比如variation的值是3，即0000 0000 0000 0011

                    对应的是0×0001和0×0002：
                    0000 0000 0000 0001
                    0000 0000 0000 0010
                """

                # 读取数据 HatBoxClass
                mainAST = ast.children[0]

                # 读取latex数据
                mainSlot, _ = self.makeLatex(mainAST)

                # 转成latex代码
                mainStr = ''
                if mainSlot != "":
                    mainStr = "{ %s }" % mainSlot

                """
                    variation转码
                """
                variationsMap = {}  # map[uint16]string
                variationsMap[0x0001] = "left"
                variationsMap[0x0002] = "right"
                variationsMap[0x0004] = "tvVE_UNDER"
                variationsMap[0x0008] = "harpoonup"

                # 有序循环
                variationsCode = [0x0001, 0x0002, 0x0004, 0x0008]  # []uint16{0x0001, 0x0002, 0x0004, 0x0008}

                topStr = "\\overset\\"
                for vCode in variationsCode:
                    if vCode & tmpl.variation != 0:
                        topStr = topStr + variationsMap[vCode]

                # 如果variationCode小于8，则一定不是harpoon,那么默认就使用arrow
                if tmpl.variation < 8:
                    topStr = topStr + "arrow"
                """
                    variation转码 END
                """

                # 组成整体公式
                tmplStr = "%s %s" % (topStr, mainStr)
                buf += tmplStr

                return buf, None
            elif tmpl.selector == SelectorType.tmHAT:
                # 读取数据 HatBoxClass
                mainAST = ast.children[0]
                topAST = ast.children[1]

                # 读取latex数据
                mainSlot, _ = self.makeLatex(mainAST)
                topSlot, _ = self.makeLatex(topAST)

                # 转成latex代码
                mainStr = ''
                topStr = ''
                if mainSlot != "":
                    mainStr = "{ %s }" % mainSlot
                if topSlot != "":
                    topStr = " %s " % topSlot

                # 组成整体公式
                tmplStr = "%s %s" % (topStr, mainStr)
                buf += tmplStr

                return buf, None
            elif tmpl.selector == SelectorType.tmARC:
                # 读取数据 HatBoxClass
                mainAST = ast.children[0]
                topAST = ast.children[1]

                # 读取latex数据
                mainSlot, _ = self.makeLatex(mainAST)
                topSlot, _ = self.makeLatex(topAST)

                # 转成latex代码
                mainStr = ''
                topStr = ''
                if mainSlot != "":
                    mainStr = "{ %s }" % mainSlot
                if topSlot != "":
                    topStr = "\\overset %s" % topSlot

                # 组成整体公式
                tmplStr = "%s %s" % (topStr, mainStr)
                buf += tmplStr

                return buf, None
            elif tmpl.selector == SelectorType.tmINTEG:
                # 积分符号处理
                mainSlot = ''
                lowerSlot = ''
                upperSlot = ''
                idx = 0
                for astData in ast.children:
                    if idx == 0:
                        mainSlot, _ = self.makeLatex(astData)
                    elif idx == 1:
                        lowerSlot, _ = self.makeLatex(astData)
                    elif idx == 2:
                        upperSlot, _ = self.makeLatex(astData)
                    idx += 1

                # 根据variation决定积分类型
                integral_symbol = '\\int'
                if tmpl.variation & 0x0002:  # tvINT_2 - 双重积分
                    integral_symbol = '\\iint'
                elif tmpl.variation & 0x0003:  # tvINT_3 - 三重积分
                    integral_symbol = '\\iiint'
                elif tmpl.variation & 0x0004:  # tvINT_LOOP - 环积分
                    integral_symbol = '\\oint'

                # 转成latex代码
                mainStr = f'{{ {mainSlot} }}' if mainSlot else ''
                lowerStr = f'_{{{lowerSlot}}}' if lowerSlot else ''
                upperStr = f'^{{{upperSlot}}}' if upperSlot else ''

                tmplStr = f'{integral_symbol}{lowerStr}{upperStr} {mainStr}'
                buf += tmplStr
                return buf, None
            elif tmpl.selector == SelectorType.tmPROD:
                # 乘积符号处理
                mainSlot = ''
                lowerSlot = ''
                upperSlot = ''
                idx = 0
                for astData in ast.children:
                    if idx == 0:
                        mainSlot, _ = self.makeLatex(astData)
                    elif idx == 1:
                        lowerSlot, _ = self.makeLatex(astData)
                    elif idx == 2:
                        upperSlot, _ = self.makeLatex(astData)
                    idx += 1

                # 转成latex代码
                mainStr = f'{{ {mainSlot} }}' if mainSlot else ''
                lowerStr = f'\\limits_{{ {lowerSlot} }}' if lowerSlot else ''
                upperStr = f'^{{ {upperSlot} }}' if upperSlot else ''

                tmplStr = f'\\prod {lowerStr}{upperStr} {mainStr}'
                buf += tmplStr
                return buf, None
            elif tmpl.selector == SelectorType.tmTILDE:
                # 波浪号装饰
                mainAST = ast.children[0] if ast.children else None
                if mainAST:
                    mainSlot, _ = self.makeLatex(mainAST)
                    buf += f'\\tilde{{ {mainSlot} }}'
                return buf, None
            elif tmpl.selector == SelectorType.tmFLOOR:
                # 向下取整 (floor brackets)
                mainAST = ast.children[0] if ast.children else None
                if mainAST:
                    mainSlot, _ = self.makeLatex(mainAST)
                    buf += f'\\lfloor {mainSlot} \\rfloor'
                return buf, None
            elif tmpl.selector == SelectorType.tmCEILING:
                # 向上取整 (ceiling brackets)
                mainAST = ast.children[0] if ast.children else None
                if mainAST:
                    mainSlot, _ = self.makeLatex(mainAST)
                    buf += f'\\lceil {mainSlot} \\rceil'
                return buf, None
            elif tmpl.selector == SelectorType.tmDBAR:
                # 双竖线 (double vertical bars)
                mainAST = ast.children[0] if ast.children else None
                if mainAST:
                    mainSlot, _ = self.makeLatex(mainAST)
                    buf += f'\\| {mainSlot} \\|'
                return buf, None
            elif tmpl.selector == SelectorType.tmINTOP:
                # 积分样式大型操作符处理 (big integral-style operators)
                mainSlot = ''
                lowerSlot = ''
                upperSlot = ''
                operatorChar = ''

                # 解析子对象 - BigOpBoxClass 的顺序：main slot, lower slot, upper slot, large operator character
                idx = 0
                for astData in ast.children:
                    if idx == 0:  # main slot (被操作的表达式)
                        mainSlot, _ = self.makeLatex(astData)
                    elif idx == 1:  # lower slot (下限)
                        lowerSlot, _ = self.makeLatex(astData)
                    elif idx == 2:  # upper slot (上限)
                        upperSlot, _ = self.makeLatex(astData)
                    elif idx == 3:  # large operator character (大型操作符字符)
                        operatorChar, _ = self.makeLatex(astData)
                    idx += 1

                # 根据variation决定限制的位置
                # 0: tvUINTOP - upper limit
                # 1: tvLINTOP - lower limit
                # 2: tvBINTOP - both upper and lower limits

                # 积分样式意味着限制放在操作符右侧，而非上下方
                # 如果没有操作符字符，使用默认的积分样式操作符
                if not operatorChar:
                    operatorChar = '\\bigodot'  # 默认使用大圆点操作符

                # 转成latex代码
                mainStr = f'{{ {mainSlot} }}' if mainSlot else ''
                lowerStr = f'_{{{lowerSlot}}}' if lowerSlot else ''
                upperStr = f'^{{{upperSlot}}}' if upperSlot else ''

                # 积分样式布局：操作符 + 上下标 + 被操作表达式
                tmplStr = f'{operatorChar}{lowerStr}{upperStr} {mainStr}'
                buf += tmplStr
                return buf, None
            else:
                # self.Valid = False
                buf = "latex tmpl not implement"  # 设置一个特殊字符方便在论文中定位
                logger.warning('MTEF.makeLatex:TMPL NOT IMPLEMENT, %s, %s', tmpl.selector, tmpl.variation)
            for _ast in ast.children:
                _latex, _ = self.makeLatex(_ast)
                buf += _latex
            return buf, None
        elif ast.tag == RecordType.PILE:
            idx = 0
            for _ast in ast.children:
                _latex, _ = self.makeLatex(_ast)

                # 多个line字符串数据以 \\ 分割
                if idx > 0:
                    buf + " \\\\ "

                buf += _latex
                idx += 1
            return buf, None
        elif ast.tag == RecordType.MATRIX:
            matrixCol = int(ast.value.cols)
            idx = 0
            for _ast in ast.children:
                _latex, _ = self.makeLatex(_ast)

                if idx == 0:
                    buf += " \\begin{array} {} "
                    continue

                buf += _latex

                if idx % matrixCol == 0:
                    buf += " \\\\ "
                else:
                    buf += " & "
                idx += 1

            buf += " \\end{array} "
            return buf, None
        elif ast.tag == RecordType.LINE:
            for _ast in ast.children:
                _latex, _ = self.makeLatex(_ast)
                buf += _latex
            return buf, None
        elif ast.tag == RecordType.EMBELL:
            embellType = ast.value.embellType
            embellMapping = self.getEmbellMapping(is_v3=False)

            embellStr = embellMapping.get(embellType, "")
            if not embellStr:
                logger.warning('MTEF.makeLatex:not implement embell: %s', embellType)
                embellStr = ""
            else:
                # 对于需要参数的装饰符号，添加空格分隔
                if not embellStr.startswith("'"):
                    embellStr = f" {embellStr} "

            buf += embellStr
            return buf, None

        return '', None

    def makeLatexV3(self, ast):
        """
        为 MTEF v3 版本生成 LaTeX 代码
        处理 v3 特有的模板选择器和变体代码
        """
        if ast is None:
            return '', None

        if ast.tag == RecordTypeV3.END:
            return '', None

        elif ast.tag == RecordTypeV3.LINE:
            if ast.value and ast.value.null:
                return '', None

            latex_parts = []
            if ast.children:
                for child in ast.children:
                    child_latex, err = self.makeLatexV3(child)
                    if err:
                        return '', err
                    if child_latex:
                        latex_parts.append(child_latex)
            return ''.join(latex_parts), None

        elif ast.tag == RecordTypeV3.CHAR:
            if not ast.value:
                return '', None

            # 处理字符的 typeface 和 mtcode
            char = ast.value
            typeface = char.typeface - 128 if char.typeface >= 128 else char.typeface
            mtcode = char.mtcode

            # 根据 MTCode 转换为相应的 LaTeX 字符
            if mtcode == 0x0050:  # 'P'
                return 'P', None
            elif mtcode == 0x0068:  # 'h'
                return 'h', None
            elif mtcode == 0x003D:  # '='
                return '=', None
            elif mtcode == 0x0030:  # '0'
                return '0', None
            elif mtcode >= 0x0041 and mtcode <= 0x005A:  # A-Z
                return chr(mtcode), None
            elif mtcode >= 0x0061 and mtcode <= 0x007A:  # a-z
                return chr(mtcode), None
            elif mtcode >= 0x0030 and mtcode <= 0x0039:  # 0-9
                return chr(mtcode), None
            else:
                # 其他字符，尝试直接转换
                try:
                    return chr(mtcode), None
                except:
                    return f'\\text{{{mtcode}}}', None

        elif ast.tag == RecordTypeV3.TMPL:
            if not ast.value:
                return '', None

            tmpl = ast.value
            selector = tmpl.selector
            variation = tmpl.variation

            # 处理各种模板类型
            if selector == SelectorTypeV3.tmFRACT:  # 分数
                if len(ast.children) >= 2:
                    numerator, err = self.makeLatexV3(ast.children[0])
                    if err:
                        return '', err
                    denominator, err = self.makeLatexV3(ast.children[1])
                    if err:
                        return '', err
                    return f'\\frac{{{numerator}}}{{{denominator}}}', None

            elif selector == SelectorTypeV3.tmSINT:  # 单积分
                integral_symbol = '\\int'

                # 在当前 AST 结构中，所有子节点都在一个列表中
                # 需要根据实际内容智能分组
                main_parts = []

                if ast.children:
                    for child in ast.children:
                        child_latex, err = self.makeLatexV3(child)
                        if err:
                            return '', err
                        if child_latex:
                            main_parts.append(child_latex)

                main_slot = ''.join(main_parts)

                # 根据 variation 决定积分的样式
                if variation == 0:  # tvNSINT - no limits
                    return f'{integral_symbol} {main_slot}', None
                elif variation == 1:  # tvLSINT - lower limit only
                    # 对于这个变体，我们暂时简化处理
                    return f'{integral_symbol} {main_slot}', None
                elif variation == 2:  # tvBSINT - both limits
                    return f'{integral_symbol} {main_slot}', None
                elif variation == 3:  # tvNCINT - contour, no limits
                    return f'\\oint {main_slot}', None
                elif variation == 4:  # tvLCINT - contour, lower limit only
                    return f'\\oint {main_slot}', None

            elif selector == SelectorTypeV3.tmSCRIPT:  # 上下标
                # 从调试中发现，上标模板通常有两个子节点
                # 第一个子节点通常为空，第二个子节点包含实际内容
                superscript_content = ''
                subscript_content = ''

                # 寻找非空的子节点作为上标或下标内容
                for child in ast.children:
                    child_latex, err = self.makeLatexV3(child)
                    if err:
                        return '', err
                    if child_latex.strip():  # 非空内容
                        if variation == 0:  # tvSUPER - superscript
                            superscript_content = child_latex
                        elif variation == 1:  # tvSUB - subscript
                            subscript_content = child_latex
                        break  # 找到第一个非空内容就停止

                if variation == 0:  # tvSUPER - superscript
                    return f'^{{{superscript_content}}}', None
                elif variation == 1:  # tvSUB - subscript
                    return f'_{{{subscript_content}}}', None
                elif variation == 2:  # tvSUBSUP - both
                    if len(ast.children) >= 2:
                        subscript, err = self.makeLatexV3(ast.children[0])
                        if err:
                            return '', err
                        superscript, err = self.makeLatexV3(ast.children[1])
                        if err:
                            return '', err
                        return f'_{{{subscript}}}^{{{superscript}}}', None

            elif selector == SelectorTypeV3.tmROOT:  # 根号
                if variation == 0:  # tvSQROOT - square root
                    if ast.children:
                        radicand, err = self.makeLatexV3(ast.children[0])
                        if err:
                            return '', err
                        return f'\\sqrt{{{radicand}}}', None
                elif variation == 1:  # tvNTHROOT - nth root
                    if len(ast.children) >= 2:
                        main_slot, err = self.makeLatexV3(ast.children[0])
                        if err:
                            return '', err
                        nth_slot, err = self.makeLatexV3(ast.children[1])
                        if err:
                            return '', err
                        return f'\\sqrt[{nth_slot}]{{{main_slot}}}', None

            elif selector == SelectorTypeV3.tmPAREN:  # 括号
                if ast.children:
                    content, err = self.makeLatexV3(ast.children[0])
                    if err:
                        return '', err

                    if variation == 0:  # tvBPAREN - both left and right
                        return f'\\left({content}\\right)', None
                    elif variation == 1:  # tvLPAREN - left only
                        return f'\\left({content}\\right.', None
                    elif variation == 2:  # tvRPAREN - right only
                        return f'\\left.{content}\\right)', None

            elif selector == SelectorTypeV3.tmBRACK:  # 方括号
                if ast.children:
                    content, err = self.makeLatexV3(ast.children[0])
                    if err:
                        return '', err

                    if variation == 0:  # tvBBRACK - both left and right
                        return f'\\left[{content}\\right]', None
                    elif variation == 1:  # tvLBRACK - left only
                        return f'\\left[{content}\\right.', None
                    elif variation == 2:  # tvRBRACK - right only
                        return f'\\left.{content}\\right]', None

            elif selector == SelectorTypeV3.tmBRACE:  # 花括号
                if ast.children:
                    content, err = self.makeLatexV3(ast.children[0])
                    if err:
                        return '', err

                    if variation == 0:  # tvBBRACE - both left and right
                        return f'\\left\\{{{content}\\right\\}}', None
                    elif variation == 1:  # tvLBRACE - left only
                        return f'\\left\\{{{content}\\right.', None
                    elif variation == 2:  # tvRBRACE - right only
                        return f'\\left.{content}\\right\\}}', None

            elif selector == SelectorTypeV3.tmSUM:  # 求和
                sum_symbol = '\\sum'
                main_slot = ''
                lower_slot = ''
                upper_slot = ''

                if ast.children:
                    if len(ast.children) >= 1:
                        main_slot, err = self.makeLatexV3(ast.children[0])
                        if err:
                            return '', err
                    if len(ast.children) >= 2:
                        upper_slot, err = self.makeLatexV3(ast.children[1])
                        if err:
                            return '', err
                    if len(ast.children) >= 3:
                        lower_slot, err = self.makeLatexV3(ast.children[2])
                        if err:
                            return '', err

                if variation == 0:  # tvLSUM - lower only
                    return f'{sum_symbol}_{{{lower_slot}}} {main_slot}', None
                elif variation == 1:  # tvBSUM - both upper and lower limits
                    return f'{sum_symbol}_{{{lower_slot}}}^{{{upper_slot}}} {main_slot}', None
                elif variation == 2:  # tvNSUM - no limits
                    return f'{sum_symbol} {main_slot}', None

            elif selector == SelectorTypeV3.tmPROD:  # 乘积
                prod_symbol = '\\prod'
                main_slot = ''
                lower_slot = ''
                upper_slot = ''

                if ast.children:
                    if len(ast.children) >= 1:
                        main_slot, err = self.makeLatexV3(ast.children[0])
                        if err:
                            return '', err
                    if len(ast.children) >= 2:
                        upper_slot, err = self.makeLatexV3(ast.children[1])
                        if err:
                            return '', err
                    if len(ast.children) >= 3:
                        lower_slot, err = self.makeLatexV3(ast.children[2])
                        if err:
                            return '', err

                if variation == 0:  # tvLPROD - lower only
                    return f'{prod_symbol}_{{{lower_slot}}} {main_slot}', None
                elif variation == 1:  # tvBPROD - both upper and lower limits
                    return f'{prod_symbol}_{{{lower_slot}}}^{{{upper_slot}}} {main_slot}', None
                elif variation == 2:  # tvNPROD - no limits
                    return f'{prod_symbol} {main_slot}', None
            elif selector == SelectorTypeV3.tmLSCRIPT:  # 左上标和左下标
                # 左上标和左下标的处理
                if variation == 0:  # tvLSUPER - 左上标
                    if ast.children:
                        superscript_content = ''
                        for child in ast.children:
                            child_latex, err = self.makeLatexV3(child)
                            if err:
                                return '', err
                            if child_latex.strip():
                                superscript_content = child_latex
                                break
                        return f'{{}}^{{{superscript_content}}}', None
                elif variation == 1:  # tvLSUB - 左下标
                    if ast.children:
                        subscript_content = ''
                        for child in ast.children:
                            child_latex, err = self.makeLatexV3(child)
                            if err:
                                return '', err
                            if child_latex.strip():
                                subscript_content = child_latex
                                break
                        return f"{{}}_{{{subscript_content}}}", None
                elif variation == 2:  # tvLSUBSUP - 左上标和左下标
                    if len(ast.children) >= 2:
                        subscript, err = self.makeLatexV3(ast.children[0])
                        if err:
                            return '', err
                        superscript, err = self.makeLatexV3(ast.children[1])
                        if err:
                            return '', err
                        return f'{{}}_{{{subscript}}}^{{{superscript}}}', None

            # 默认情况：处理未明确支持的模板
            latex_parts = []
            if ast.children:
                for child in ast.children:
                    child_latex, err = self.makeLatexV3(child)
                    if err:
                        return '', err
                    if child_latex:
                        latex_parts.append(child_latex)
            return ''.join(latex_parts), None

        elif ast.tag == RecordTypeV3.PILE:
            # 处理垂直堆叠（多行）
            latex_parts = []
            if ast.children:
                for i, child in enumerate(ast.children):
                    child_latex, err = self.makeLatexV3(child)
                    if err:
                        return '', err
                    if child_latex:
                        latex_parts.append(child_latex)
                        if i < len(ast.children) - 1:
                            latex_parts.append(' \\\\ ')

            if len(latex_parts) > 1:
                return f'\\begin{{aligned}} {" ".join(latex_parts)} \\end{{aligned}}', None
            else:
                return ''.join(latex_parts), None

        elif ast.tag == RecordTypeV3.MATRIX:
            # 处理矩阵
            if not ast.value:
                return '', None

            matrix = ast.value
            rows = matrix.rows
            cols = matrix.cols

            latex_parts = ['\\begin{pmatrix}']

            if ast.children:
                for i in range(rows):
                    row_parts = []
                    for j in range(cols):
                        cell_index = i * cols + j
                        if cell_index < len(ast.children):
                            cell_latex, err = self.makeLatexV3(
                                ast.children[cell_index])
                            if err:
                                return '', err
                            row_parts.append(cell_latex)
                        else:
                            row_parts.append('')

                    latex_parts.append(' & '.join(row_parts))
                    if i < rows - 1:
                        latex_parts.append(' \\\\ ')

            latex_parts.append('\\end{pmatrix}')
            return ''.join(latex_parts), None

        elif ast.tag == RecordTypeV3.EMBELL:
            # 处理装饰（帽子、点等）
            if not ast.value:
                return '', None

            embell = ast.value
            embell_type = embell.embellType
            embellMapping = self.getEmbellMapping(is_v3=True)

            embellStr = embellMapping.get(embell_type, "")
            if not embellStr:
                logger.warning('MTEF.makeLatexV3:not implement embell: %s', embell_type)
                # 如果没有找到对应的装饰，直接返回子内容
                if ast.children:
                    base_latex, err = self.makeLatexV3(ast.children[0])
                    if err:
                        return '', err
                    return base_latex, None
                return '', None

            if ast.children:
                base_latex, err = self.makeLatexV3(ast.children[0])
                if err:
                    return '', err

                # 对于撇号类装饰，直接追加到内容后面
                if embellStr.startswith("'"):
                    return f'{base_latex}{embellStr}', None
                else:
                    # 对于其他装饰，使用花括号包围
                    return f'{embellStr}{{{base_latex}}}', None

            return '', None

        elif ast.tag in (RecordTypeV3.FULL, RecordTypeV3.SUB, RecordTypeV3.SUB2,
                         RecordTypeV3.SYM, RecordTypeV3.SUBSYM):
            # 大小控制记录，不直接生成 LaTeX
            return '', None
        else:
            # 处理其他节点类型，递归处理子节点
            latex_parts = []
            if ast.children:
                for child in ast.children:
                    child_latex, err = self.makeLatexV3(child)
                    if err:
                        return '', err
                    if child_latex:
                        latex_parts.append(child_latex)
            return ''.join(latex_parts), None

    @classmethod
    def OpenBytes(cls, bts):
        return cls.Open(BytesIO(bts))

    @classmethod
    def Open(cls, reader):
        ole, err = Ole.Open(reader)
        if err is not None:
            logger.error(err)

        dir, err = ole.ListDir()
        if err is not None:
            logger.error(err)

        for file in dir:
            if 'Equation Native' == file.Name():
                root = dir[0]
                reader = ole.OpenFile(file, root)

                hdrBuffer = reader.read(oleCbHdr)
                if hdrBuffer is not None and len(hdrBuffer) == oleCbHdr:
                    hdrReader = BytesIO(hdrBuffer)
                    cbHdr = 0  # uint16
                    cbSize = 0  # uint32

                    cbHdr = Helper.bytes2int(hdrReader.read(2))
                    if cbHdr is None or cbHdr != oleCbHdr:
                        return None, 'MTEF.Open: read byte error'

                    # ignore 'version: u32' and 'cf: u16'
                    hdrReader.seek(4+2, 1)  # io.SeekCurrent
                    cbSize = Helper.bytes2int(hdrReader.read(4))

                    # body from 'cbHdr' to 'cbHdr + cbSize'
                    reader.seek(cbHdr, 0)  # io.SeekStart
                    real_size = file.Size - oleCbHdr
                    eqnBody = reader.read(real_size)
                    eqn = MTEF()
                    eqn.reader = BytesIO(eqnBody)

                    eqn.readRecord()
                    eqn.makeAST()
                    return eqn, None

                return None, 'MTEF.Open: read byte error'

        return None, err

    def getEmbellMapping(self, is_v3=False):
        """
        获取装饰类型到 LaTeX 符号的映射表
        is_v3: True表示V3版本，False表示V5版本
        """
        if is_v3:
            # MTEF V3 版本的装饰映射
            return {
                EmbellTypeV3.embDOT: "\\dot",
                EmbellTypeV3.embDDOT: "\\ddot",
                EmbellTypeV3.embTDOT: "\\dddot",
                EmbellTypeV3.embPRIME: "'",
                EmbellTypeV3.embDPRIME: "''",
                EmbellTypeV3.embTPRIME: "'''",
                EmbellTypeV3.embBPRIME: "^\\backprime",
                EmbellTypeV3.embTILDE: "\\tilde",
                EmbellTypeV3.embHAT: "\\hat",
                EmbellTypeV3.embNOT: "\\not",
                EmbellTypeV3.embRARROW: "\\overrightarrow",
                EmbellTypeV3.embLARROW: "\\overleftarrow",
                EmbellTypeV3.embBARROW: "\\overleftrightarrow",
                EmbellTypeV3.embR1ARROW: "\\overrightarrow",  # 单倒钩箭头暂用普通箭头
                EmbellTypeV3.embL1ARROW: "\\overleftarrow",   # 单倒钩箭头暂用普通箭头
                EmbellTypeV3.embMBAR: "\\overline",  # 中高度横线暂用上划线
                EmbellTypeV3.embOBAR: "\\overline",
                EmbellTypeV3.embFROWN: "\\frown",
                EmbellTypeV3.embSMILE: "\\smile",
            }
        else:
            # MTEF V5 版本的装饰映射
            return {
                EmbellType.emb1DOT: "\\dot",
                EmbellType.emb2DOT: "\\ddot",
                EmbellType.emb3DOT: "\\dddot",
                EmbellType.emb1PRIME: "'",
                EmbellType.emb2PRIME: "''",
                EmbellType.emb3PRIME: "'''",
                EmbellType.embBPRIME: "^\\backprime",
                EmbellType.embTILDE: "\\tilde",
                EmbellType.embHAT: "\\hat",
                EmbellType.embNOT: "\\not",
                EmbellType.embRARROW: "\\overrightarrow",
                EmbellType.embLARROW: "\\overleftarrow",
                EmbellType.embBARROW: "\\overleftrightarrow",
                EmbellType.embR1ARROW: "\\overrightarrow",  # 单倒钩箭头暂用普通箭头
                EmbellType.embL1ARROW: "\\overleftarrow",   # 单倒钩箭头暂用普通箭头
                EmbellType.embMBAR: "\\overline",  # 中高度横线暂用上划线
                EmbellType.embOBAR: "\\overline",
                EmbellType.embFROWN: "\\frown",
                EmbellType.embSMILE: "\\smile",
                EmbellType.embX_BARS: "\\cancel",  # 双对角线暂用取消线
                EmbellType.embUP_BAR: "\\nearrow",  # 左下到右上对角线暂用箭头
                EmbellType.embDOWN_BAR: "\\searrow",  # 左上到右下对角线暂用箭头
                EmbellType.emb4DOT: "\\ddddot",  # 四点装饰需要特殊宏包
                EmbellType.embU_1DOT: "\\underdot",  # 下单点
                EmbellType.embU_2DOT: "\\underddot",  # 下双点
                EmbellType.embU_3DOT: "\\underdddot",  # 下三点
                EmbellType.embU_4DOT: "\\underddddot",  # 下四点
                EmbellType.embU_BAR: "\\underline",
                EmbellType.embU_TILDE: "\\undertilde",  # 下波浪线
                EmbellType.embU_FROWN: "\\underfrown",  # 下弧线
                EmbellType.embU_SMILE: "\\undersmile",  # 下弧线
                EmbellType.embU_RARROW: "\\underrightarrow",  # 下右箭头
                EmbellType.embU_LARROW: "\\underleftarrow",  # 下左箭头
                EmbellType.embU_BARROW: "\\underleftrightarrow",  # 下双向箭头
                EmbellType.embU_R1ARROW: "\\underrightarrow",  # 下右单倒钩箭头
                EmbellType.embU_L1ARROW: "\\underleftarrow",   # 下左单倒钩箭头
            }
