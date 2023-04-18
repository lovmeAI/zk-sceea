# coding: utf-8
# +-------------------------------------------------------------------
# | Project     : zk-sceea
# | Author      : 晴天雨露
# | QQ/Email    : 184566608<qingyueheji@qq.com>
# | Time        : 2022/8/17 22:52
# | Describe    : utils
# +-------------------------------------------------------------------
def bk_done(e, jf_log, addr1):
    """
    报考完毕并且已缴费
    :param e:
    :param jf_log:
    :param addr1:
    :return:
    """
    location = e.xpath("//table[@id='tablePlace']//td[3]/text()")
    if len(location) > 0:
        location = location[0]
    _yx_course_list = e.xpath("//table[@id='tableYXCourse']/tbody/tr/td[3]/text()")
    courses = '、'.join(_yx_course_list[::2])
    return f"{jf_log}完成报考项目: {courses}，报考位置({addr1}/{location})"

