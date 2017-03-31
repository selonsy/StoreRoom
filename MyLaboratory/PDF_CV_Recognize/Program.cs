using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using org.apache.pdfbox.pdmodel;
using org.apache.pdfbox.util;
using System.IO;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System.Text.RegularExpressions;

using MsWord = Microsoft.Office.Interop.Word;

namespace PDF_CV_Recognize
{
    class Program
    {
        #region 外部资源

        #region 中国的56个民族

        public string[] Nations = new string[] {
            "壮族","藏族","裕固族","彝族","瑶族","锡伯族","乌孜别克族","维吾尔族","佤族","土家族","土族","塔塔尔族",
            "塔吉克族","水族","畲族","撒拉族","羌族","普米族","怒族","纳西族","仫佬族","苗族","蒙古族","门巴族",
            "毛南族","满族","珞巴族","僳僳族","黎族","拉祜族","柯尔克孜族","景颇族","京族","基诺族","回族","赫哲族",
            "哈萨克族","哈尼族","仡佬族","高山族","鄂温克族","俄罗斯族","鄂伦春族","独龙族","东乡族","侗族","德昂族",
            "傣族","达斡尔族","朝鲜族","布依族","布朗族","保安族","白族","阿昌族",
            "汉族"
        };

        #endregion

        #region 百家姓

        public static string[] BaiJiaXing_Double = new string[] {
            "东郭", "南门", "呼延", "羊舌", "微生", "左丘",
            "万俟", "司马", "上官", "欧阳", "夏侯", "诸葛", "闻人", "东方", "赫连", "皇甫",
            "尉迟", "公羊", "澹台", "公冶", "宗政", "濮阳", "东门", "西门", "南宫", "第五",
            "淳于", "单于", "太叔", "申屠", "公孙", "仲孙", "轩辕", "令狐", "钟离", "宇文",
            "长孙", "慕容", "鲜于", "闾丘", "司徒", "司空", "亓官", "司寇", "子车", "夹谷",
            "颛孙", "端木", "巫马", "公西", "漆雕", "乐正", "壤驷", "公良", "拓跋", "梁丘",
            "宰父", "谷梁", "段干", "百里"
        };

        //可能需要按照人数进行排序
        public static string[] BaiJiaXing_Single = new string[] {
            "赵", "钱", "孙", "李", "周", "吴", "郑", "王", "冯", "陈",
            "褚", "卫", "蒋", "沈", "韩", "杨", "朱", "秦", "尤", "许",
            "何", "吕", "施", "张", "孔", "曹", "严", "华", "金", "魏",
            "陶", "姜", "戚", "谢", "邹", "喻", "柏", "水", "窦", "章",
            "云", "苏", "潘", "葛", "奚", "范", "彭", "郎", "鲁", "韦",
            "昌", "马", "苗", "凤", "花", "方", "俞", "任", "袁", "柳",
            "酆", "鲍", "史", "贺", "唐", "费", "廉", "岑", "薛", "雷",
            "倪", "汤", "滕", "殷", "罗", "毕", "郝", "邬", "安", "常",
            "乐", "于", "时", "傅", "皮", "卞", "齐", "康", "伍", "余",
            "元", "卜", "顾", "孟", "平", "黄", "和", "穆", "萧", "尹",
            "姚", "邵", "湛", "汪", "祁", "毛", "禹", "狄", "米", "贝",
            "明", "臧", "计", "伏", "成", "戴", "谈", "宋", "茅", "庞",
            "熊", "纪", "舒", "屈", "项", "祝", "董", "粱", "杜", "阮",
            "蓝", "闵", "席", "季", "麻", "强", "贾", "路", "娄", "危",
            "江", "童", "颜", "郭", "梅", "盛", "林", "刁", "钟", "徐",
            "邱", "骆", "高", "夏", "蔡", "田", "樊", "胡", "凌", "霍",
            "虞", "万", "支", "柯", "昝", "管", "卢", "莫", "经", "房",
            "裘", "缪", "干", "解", "应", "宗", "丁", "宣", "贲", "邓",
            "郁", "单", "杭", "洪", "包", "诸", "左", "石", "崔", "吉",
            "钮", "龚", "程", "嵇", "邢", "滑", "裴", "陆", "荣", "翁",
            "荀", "羊", "於", "惠", "甄", "麴", "家", "封", "芮", "羿",
            "储", "靳", "汲", "邴", "糜", "松", "井", "段", "富", "巫",
            "乌", "焦", "巴", "弓", "牧", "隗", "山", "谷", "车", "侯",
            "宓", "蓬", "全", "郗", "班", "仰", "秋", "仲", "伊", "宫",
            "宁", "仇", "栾", "暴", "甘", "钭", "厉", "戎", "祖", "武",
            "符", "刘", "景", "詹", "束", "龙", "叶", "幸", "司", "韶",
            "郜", "黎", "蓟", "薄", "印", "宿", "白", "怀", "蒲", "邰",
            "从", "鄂", "索", "咸", "籍", "赖", "卓", "蔺", "屠", "蒙",
            "池", "乔", "阴", "欎", "胥", "能", "苍", "双", "闻", "莘",
            "党", "翟", "谭", "贡", "劳", "逄", "姬", "申", "扶", "堵",
            "冉", "宰", "郦", "雍", "舄", "璩", "桑", "桂", "濮", "牛",
            "寿", "通", "边", "扈", "燕", "冀", "郏", "浦", "尚", "农",
            "温", "别", "庄", "晏", "柴", "瞿", "阎", "充", "慕", "连",
            "茹", "习", "宦", "艾", "鱼", "容", "向", "古", "易", "慎",
            "戈", "廖", "庾", "终", "暨", "居", "衡", "步", "都", "耿",
            "满", "弘", "匡", "国", "文", "寇", "广", "禄", "阙", "东",
            "殴", "殳", "沃", "利", "蔚", "越", "夔", "隆", "师", "巩",
            "厍", "聂", "晁", "勾", "敖", "融", "冷", "訾", "辛", "阚",
            "那", "简", "饶", "空", "曾", "毋", "沙", "乜", "养", "鞠",
            "须", "丰", "巢", "关", "蒯", "相", "查", "後", "荆", "红",
            "游", "竺", "权", "逯", "盖", "益", "桓", "公", "墨", "哈",
            "谯", "笪", "年", "爱", "阳", "佟", "商", "帅", "佘", "佴",
            "仉", "督", "归", "海", "伯", "赏", "岳", "楚", "缑", "亢",
            "况", "后", "有", "琴", "言", "福", "晋", "牟", "闫", "法",
            "汝", "鄢", "涂", "钦"
        };

        #endregion

        #endregion

        public static void Recongnize_PDF_Resume(string filepath)
        {
            PDDocument doc = null;
            try
            {
                doc = PDDocument.load(filepath);
                PDFTextStripper stripper = new PDFTextStripper();
                //读取PDF中的信息
                string text = stripper.getText(doc);

                //string word_str = GetContentFromWord("./file/智联招聘-test-word.doc");
                //输出内容到txt文件
                string name = filepath.Replace("/","").Replace(".","");
                File.WriteAllText("./file/output/" + name + ".txt", text);

                //解析PDF内容,并填充到简历信息中
                Resume resume = GetResumeInfo(text);

                Console.WriteLine("\nBegin of {0}\n", filepath);

                #region DOS显示读出的显示

                //将读出的信息输出到控制台中,以键值对的形式查看
                string str_resume = JsonConvert.SerializeObject(resume);
                JsonTextReader reader = new JsonTextReader(new StringReader(str_resume));
                int i = 1;
                string str = "";
                while (reader.Read())
                {
                    if (reader.TokenType.ToString() == "PropertyName")
                    {
                        str = str + reader.Value + ":";
                    }
                    else
                    {
                        if (reader.Value != null)
                        {
                            i++;
                            str += reader.Value;
                        }
                        else
                        {
                            str = "";
                            i = 1;
                        }
                    }
                    if (i == 2)
                    {
                        Console.WriteLine(str);
                        i = 1;
                        str = "";
                    }
                }

                #endregion
                Console.WriteLine("\nEnd of {0}\n",filepath);
            }
            finally
            {
                if (doc != null)
                {
                    doc.close();
                }
            }
        }

        static void Main(string[] args)
        {

            DirectoryInfo dir = new DirectoryInfo(@"D:\00MyWorkSpace\99MyGitHub\MyProject\MyLaboratory\PDF_CV_Recognize\bin\Debug\file\HR_Resume");

            foreach (var item in dir.GetFiles())
            {
                string path = "./file/HR_Resume/" + item.Name;
                Recongnize_PDF_Resume(path);
            }

            Console.ReadKey();
        }

        public static Resume GetResumeInfo(string resume_str)
        {
            Resume resume = new Resume();

            #region 姓名

            string name = string.Empty;

            name = find_expected(resume_str, new string[] { "姓名", "名字" }, 4);

            if (string.IsNullOrEmpty(name))
            {
                //进行第二次识别,正则匹配中文姓名,先匹配单姓的               
                for (int i = 0; i < BaiJiaXing_Single.Length; i++)
                {
                    //定位
                    if (BaiJiaXing_Single[i] == "广")
                    {
                        int a = 1;
                    }
                    string _reg = @"[\s^](" + BaiJiaXing_Single[i] + @"[\u4e00-\u9fa5]{1,2}?)\s";
                    Regex name_reg = new Regex(_reg);

                    string[] resume_strs = resume_str.Split(new string[] { "\r\n" }, StringSplitOptions.RemoveEmptyEntries);

                    foreach (var item in resume_strs)
                    {
                        if (name_reg.IsMatch(item))
                        {
                            Match mat = name_reg.Match(item);
                            name = mat.Groups[0].ToString().Trim();
                        }
                        if (!string.IsNullOrEmpty(name)) { goto End; }
                    }                   
                }
                End:
                if (string.IsNullOrEmpty(name))
                {
                    //匹配复姓的
                    for (int i = 0; i < BaiJiaXing_Double.Length; i++)
                    {
                        string _reg = @"\s(" + BaiJiaXing_Double[i] + @"[\u4e00-\u9fa5]{1,3}?)\s";
                        Regex name_reg = new Regex(_reg);
                        if (name_reg.IsMatch(resume_str))
                        {
                            Match mat = name_reg.Match(resume_str);
                            name = mat.Groups[0].ToString().Trim();
                        }
                        if (!string.IsNullOrEmpty(name)) break;
                    }
                }
            }

            //赋值
            resume.Name = name;

            #endregion

            #region 性别

            string sex = string.Empty;

            //正则匹配性别
            Regex sex_reg = new Regex(@"\s[\u7537\u5973]\s");
            if (sex_reg.IsMatch(resume_str))
            {
                Match mat = sex_reg.Match(resume_str);
                sex = mat.Groups[0].ToString().Trim();
            }
            if (string.IsNullOrEmpty(sex))
            {
                sex = find_expected(resume_str, new string[] { "性别" }, 1, new string[] { "男", "女" });
            }

            //赋值
            resume.Sex = sex;

            #endregion

            #region 联系电话

            string phone = string.Empty;
            //正则匹配手机或者座机号码
            Regex phone_reg = new Regex(@"\d{3}-\d{8}|\d{4}-\d{7}|[1][3-8]\d{9}");
            if (phone_reg.IsMatch(resume_str))
            {
                Match mat = phone_reg.Match(resume_str);
                phone = mat.Groups[0].ToString().Trim();
            }
            if (string.IsNullOrEmpty(phone))
            {
                //进行第二次识别,根据关键词查找
                phone = find_expected(resume_str, new string[] { "联系电话", "手机", "电话" }, 11);
            }

            //赋值
            resume.Phone = phone;

            #endregion

            #region 邮箱

            string email = string.Empty;
            //正则匹配邮箱
            Regex email_reg = new Regex(@"[a-zA-Z0-9_-]+@[a-zA-Z0-9_-]+(\.[a-zA-Z0-9_-]+)+");
            if (email_reg.IsMatch(resume_str))
            {
                Match mat = email_reg.Match(resume_str);
                email = mat.Groups[0].ToString().Trim();
            }
            if (string.IsNullOrEmpty(email))
            {
                //进行第二次识别,根据关键词查找     
                email = find_expected(resume_str, new string[] { "邮箱", "E-MAIL", "E-mail", "e-mail", "Email" }, 30);
            }

            //赋值
            resume.Email = email;

            #endregion

            #region 身份证号码

            string idcard = string.Empty;
            //正则匹配
            Regex idcard_reg = new Regex(@"\d{18}|\d{17}[Xx]");
            if (idcard_reg.IsMatch(resume_str))
            {
                Match mat = idcard_reg.Match(resume_str);
                idcard = mat.Groups[0].ToString().Trim();
            }
            if (string.IsNullOrEmpty(idcard))
            {
                //进行第二次识别,根据关键词查找
                idcard = find_expected(resume_str, new string[] { "身份证号码", "身份证" }, 18);
            }

            //赋值
            resume.IDCard = idcard;

            #endregion

            #region 民族

            string nation = string.Empty;
            //正则匹配
            Regex nation_reg = new Regex(@"\s(?!民族)([\u4e00-\u9fa5]{1,3}族)\s");
            if (nation_reg.IsMatch(resume_str))
            {
                Match mat = nation_reg.Match(resume_str);
                nation = mat.Groups[0].ToString().Trim();
            }
            if (string.IsNullOrEmpty(nation))
            {
                //进行第二次识别,根据关键词查找
                nation = find_expected(resume_str, new string[] { "民族"}, 4);
            }

            //赋值
            resume.Nation = nation;

            #endregion

            #region 毕业学校

            string school = string.Empty;
            //正则匹配
            Regex school_reg = new Regex(@"\s(?!毕业学校|毕业学院|毕业院校|宣讲学校)([\u4e00-\u9fa5]{1,8}大学|学院|学校)\s");
            if (school_reg.IsMatch(resume_str))
            {
                Match mat = school_reg.Match(resume_str);
                school = mat.Groups[0].ToString().Trim();
            }
            if (string.IsNullOrEmpty(school))
            {
                //进行第二次识别,根据关键词查找
                school = find_expected(resume_str, new string[] { "毕业学校", "毕业院校" }, 12);
            }

            //赋值
            resume.School = school;

            #endregion

            #region 学历

            string education = string.Empty;
            //正则匹配
            Regex education_reg = new Regex(@"\s本科|硕士|博士|大专|中专\s");
            if (education_reg.IsMatch(resume_str))
            {
                Match mat = education_reg.Match(resume_str);
                education = mat.Groups[0].ToString().Trim();
            }
            if (string.IsNullOrEmpty(education))
            {
                //进行第二次识别,根据关键词查找
                education = find_expected(resume_str, new string[] { "最高学历", "学历" }, 2);
            }

            //赋值
            resume.Education = education;

            #endregion

            #region 专业

            string major = string.Empty;
            if (string.IsNullOrEmpty(major))
            {
                //进行第二次识别,根据关键词查找
                major = find_expected(resume_str, new string[] { "专业" }, 10);
            }

            //赋值
            resume.Major = major;

            #endregion

            #region 籍贯

            string account = string.Empty;
            if (string.IsNullOrEmpty(account))
            {
                //进行第二次识别,根据关键词查找
                account = find_expected(resume_str, new string[] { "籍贯" , "户口所在地" }, 10);
            }

            //赋值
            resume.Account = account;

            #endregion

            return resume;
        }

        public static string find_expected(string ori, string[] tag, int length, params string[] expected)
        {
            int index = -1;
            string find_name = string.Empty;
            string result = string.Empty;
            for (int i = 0; i < tag.Length; i++)
            {
                index = ori.IndexOf(tag[i]);
                if (index != -1)
                {
                    find_name = tag[i];
                    break;
                }
            }
            if (index != -1 && !string.IsNullOrEmpty(find_name))
            {
                //找到包含指定字符串及其后面指定长度的字符串
                string find_str_ori = ori.Substring(index + find_name.Length + 1, length);

                //去除标识符以及中间的特殊字符(:,：.etc)               
                string[] str_array = find_str_ori.ToClean().Split(' ');
                result = str_array.Length > 0 && !str_array[0].ToString().IsNullOrEmpty() ? str_array[0] : "";
            }

            //验证是否是预期,不是预期的等于没找到
            if (result.IsNotNullOrEmpty() && expected.Length > 0)
            {
                bool flag = false;
                for (int i = 0; i < expected.Length; i++)
                {
                    if (result == expected[i].ToString()) flag = true;
                }
                if (!flag) result = "";
            }

            return result;
        }

        public static int find_str(string str, string[] substr, ref string email_name)
        {
            int index = -1;
            for (int i = 0; i < substr.Length; i++)
            {
                index = str.IndexOf(substr[i]);
                if (index != -1)
                {
                    email_name = substr[i];
                    break;
                }
            }
            return index;
        }

        public static string GetContentFromWord(string path)
        {
            try
            {
                MsWord.Application wordApp = new MsWord.ApplicationClass();
                wordApp.Visible = true;
                MsWord.Document wordDoc = wordApp.Documents.Open(path);
                return wordDoc.Paragraphs.Last.Range.Text;
            }
            catch (Exception ex)
            {
                return "";
            }            
        }
    }

    /// <summary>
    /// 简历信息类
    /// </summary>
    public class Resume
    {
        /// <summary>
        /// 姓名
        /// </summary>
        public string Name { get; set; }
        /// <summary>
        /// 性别
        /// </summary>
        public string Sex { get; set; }
        /// <summary>
        /// 年龄
        /// </summary>
        public int? Age { get; set; }
        /// <summary>
        /// E-mail
        /// </summary>
        public string Email { get; set; }
        /// <summary>
        /// 联系电话
        /// </summary>
        public string Phone { get; set; }
        /// <summary>
        /// 毕业院校
        /// </summary>
        public string School { get; set; }
        /// <summary>
        /// 专业
        /// </summary>
        public string Major { get; set; }
        /// <summary>
        /// 学历
        /// </summary>
        public string Education { get; set; }
        /// <summary>
        /// 求职意向
        /// </summary>
        public string DesireJob { get; set; }
        /// <summary>
        /// 身高
        /// </summary>
        public string Height { get; set; }
        /// <summary>
        /// 体重
        /// </summary>
        public string Weight { get; set; }
        /// <summary>
        /// 出生年月
        /// </summary>
        public DateTime? Birthday { get; set; }
        /// <summary>
        /// 民族
        /// </summary>
        public string Nation { get; set; }
        /// <summary>
        /// 户口所在地
        /// </summary>
        public string Account { get; set; }
        /// <summary>
        /// 院校类型(211/985/普通)
        /// </summary>
        public string SchoolType { get; set; }
        /// <summary>
        /// 毕业时间
        /// </summary>
        public DateTime? SchoolTime { get; set; }
        /// <summary>
        /// 身份证号码
        /// </summary>
        public string IDCard { get; set; }
        /// <summary>
        /// 住址
        /// </summary>
        public string Address { get; set; }
        /// <summary>
        /// 政治面貌
        /// </summary>
        public string PoliticsStatus { get; set; }
    }

    public static class Extension
    {
        public static string ToClean(this string str)
        {
            return str.Replace(":", "").Replace("：", "").Replace("\r\n"," ").Trim();
        }

        public static bool IsNullOrEmpty(this string str)
        {
            return string.IsNullOrEmpty(str);
        }

        public static bool IsNotNullOrEmpty(this string str)
        {
            return !string.IsNullOrEmpty(str);
        }
    }
}
