import re
from utils.constants import abbr_depart

# 意见区文本处理
split_ptn=r"[、|;|；|，|。|.|\s+]"
# del_ptn=r"([阅|批][办|处|示]{0,})"
del_ptn=r"([阅|批][办|处|示]{0,})|([审|会][签|核|定])"

def cleanfrac(s):
    tmp=re.sub("请|呈|(总会)|(总法)|总|(书记)|(主席)","",s)
    output=re.sub(del_ptn,"",tmp)
    output=re.sub("[()（）]","",output)
    return output


def normalizedepart(s):
    if (s=="供应链管理部") or (s=="供应链部"):
        s= abbr_depart.GY.value
    elif s=="纪监部":
        s=abbr_depart.JW.value
    elif s=="安监部":
        s=abbr_depart.A.value
    return s



def extract_names_from_niban(niban):
    niban = re.sub(r"\n","",niban)
    tmpl = re.split(split_ptn, niban)
    rawnames = [cleanfrac(s) for s in tmpl if s!=""]
    rawnames = [normalizedepart(s) for s in rawnames]
    return rawnames