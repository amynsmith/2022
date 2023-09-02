from enum import IntEnum, Enum, unique


# Because Enums are used to represent constants we recommend using UPPER_CASE names for members
# The reason for defaulting to 1 as the starting number and not 0 is that 0 is False in a boolean sense, but by default enum members all evaluate to True.

# Members of an IntEnum can be compared to integers
# However, they still can’t be compared to standard Enum enumerations

# StrEnum New in version 3.11

# Note For the majority of new code, Enum and Flag are strongly recommended, since IntEnum and IntFlag break some semantic promises of an enumeration

# options listed mainly for future evolve

@unique
class processflag(Enum):
    C = "协同"
    S = "发文"
    R = "收文"


@unique
class archivecategory_forrec(Enum):
    a = "党委"
    b = "团委"
    c = "纪委"
    d = "行政"
    e = "工会"


@unique
class choice_send(IntEnum):
    department = 0
    midleader = 1
    topleader = 2
    archive = 5
    forward = 90
    backward = 91


@unique
class choice_rec(IntEnum):
    leader = 0
    depart = 1
    archive = 2
    forward = 90
    backward = 91


topleader_keyset={"KH","VG"}

@unique
class people_leader(Enum):
    # abbr of family name
    VS = "钟XX"
    DS = "董XX"
    FU = "付XX"
    FW = "费XX"
    HJ = "韩XX"
    HL = "黄XX"
    KH = "康XX"
    QI = "齐XX"
    UI = "石XX"
    VG = "郑XX"
    VH = "张XX"
    XP = "谢XX"
    YH = "杨XX"
    YU = "喻XX"
    LG = "冷XX"
    IF = "陈XX"
    LL = "梁XX"

    @classmethod
    def describe(cls):
        string = "  ".join(list(cls.__members__.keys()))
        print(string)

    @classmethod
    def getnames(cls, inputstr):
        tmplist = inputstr.split()
        names = [cls.__members__.get(i).value if cls.__members__.__contains__(i) else i for i in tmplist]
        return names

    @classmethod
    def allinlist(cls, names):
        inputset = set(names)
        fullset = set([e.value for e in cls])
        return inputset.issubset(fullset)

    @classmethod
    def anyinlist(cls, names):
        inputset = set(names)
        fullset = set([e.value for e in cls])
        return len(inputset.intersection(fullset)) > 0

    @classmethod
    def reverselookup(cls):
        lookupdict = dict(zip([i.value[0] for i in cls], list(cls.__members__.keys())))
        lookupdict["工会主席"] = "YH"
        return lookupdict


@unique
class abbr_depart(Enum):
    A = "安监部"
    B = "办公室"
    C = "财务部"
    F = "法务部"
    GI = "工程部"
    GY = "供应链"
    G_U = "工程部设备"
    G_V = "工程部质量"
    G_W = "工程部物资"
    JU = "技术部"
    JW = "纪检部"
    R = "人资部"
    UH = "商务部"
    UI = "市场部"

    @classmethod
    def describe(cls):
        tmp = [k + i.value for k, i in cls.__members__.items()]
        string = "  ".join(tmp)
        print(string)

    @classmethod
    def reverselookup(cls):
        return dict(zip([i.value for i in cls], list(cls.__members__.keys())))


# TODO COMBINE ABBR AND PEOPLE OF DEPART INTO ONE ENUM

@unique
class people_depart(Enum):
    # abbr of depart full name
    A = "XXX"
    B = "XXX"
    C = "XXX"
    F = "XXX"
    GI = "XXX"
    GY = "XXX"
    G_U = "XXX"
    G_V = "XXX"
    G_W = "XXX"
    JU = "XXX"
    JW = "XXX"
    R = "XXX"
    UH = "XXX"
    UI = "XXX"

    @classmethod
    def describe(cls):
        string = "  ".join(list(cls.__members__.keys()))
        print(string)

    @classmethod
    def getnames(cls, inputstr):
        tmplist = inputstr.split()
        names = [cls.__members__.get(i).value if cls.__members__.__contains__(i) else i for i in tmplist]
        return names


# allow dynamic append to existing list
# def getallnames(selrange):
#     if selrange == "depart":
#         abbr_depart.describe()
#         inputstr = input(f"selrange:{selrange}\t input corresponding Capitals first...\n")
#         names = people_depart.getnames(inputstr)
#     else:
#         # selrange=="leader"
#         people_leader.describe()
#         inputstr = input(f"selrange:{selrange}\t input corresponding Capitals first...\n")
#         names = people_leader.getnames(inputstr)
#     print(names)
#     return names

# 参数个数重载 if len(*args)<=1
def getallnames(selrange, inputstr=""):
    if selrange == "depart":
        if inputstr == "":
            abbr_depart.describe()
            inputstr = input(f"selrange:{selrange}\t input corresponding Capitals first...\n")
        names = people_depart.getnames(inputstr)
    else:
        # selrange=="leader"
        if inputstr == "":
            people_leader.describe()
            inputstr = input(f"selrange:{selrange}\t input corresponding Capitals first...\n")
        names = people_leader.getnames(inputstr)
    return names


# check and forbid dynamic append to existing list of leaders
def getselrange(names):
    # for depart members, allow others not in the defined list
    # forbid exterior append to the list of existing leaders
    if people_leader.allinlist(names):
        return "leader"
    elif people_leader.anyinlist(names):
        print("Warning: mixed names")
    return "depart"


def translatenames(rawnames):
    departkeys,leaderkeys=[],[]
    unknown=[]
    reversed_departdict = abbr_depart.reverselookup()
    reversed_leaderdict = people_leader.reverselookup()
    nametodel=[]
    for n in rawnames:
        if n in reversed_leaderdict.keys():
            leaderkeys.append(reversed_leaderdict[n])
            nametodel.append(n)
            # 遍历时不删除原列表，避免 list.remove(n)
    restnames = list(set(rawnames)-set(nametodel))
    for n in restnames:
        if n in reversed_departdict.keys():
            departkeys.append(reversed_departdict[n])
        else:
            unknown.append(n)
    topkeys=list(topleader_keyset.intersection(set(leaderkeys)))
    midkeys=list(set(leaderkeys)-set(topkeys))

    departnames=[getallnames("depart", inputstr=" ".join(departkeys)) if len(departkeys)>0 else ""][0]
    midnames=[getallnames("leader", inputstr=" ".join(midkeys)) if len(midkeys)>0 else ""][0]
    topnames=[getallnames("leader", inputstr=" ".join(topkeys)) if len(topkeys)>0 else ""][0]
    return unknown,departnames,midnames,topnames


midleader_policydict = {"bf": "审批",
                        "tj": "批示",
                        "sd": "批示",
                        "xa": "批示"}


def getpolicy(processtype, selrange, templatecreator):
    if processtype == "R":
        if selrange == "leader":
            policy = "批示"
        else:
            # selrange == "depart"
            policy = "承办"
    else:
        # processtype=="S":
        # selrange == "leader":
        # limit to midleader
        policy = midleader_policydict[templatecreator]
    print(f"apply policy: {policy}")
    return policy


# avoid variable shadowing
def main():
    # abbr_depart.describe()
    # # people_depart.describe()
    # inputstr = "A B C U 123"
    # # names = people_depart.getnames(inputstr)
    # # people_leader.describe()
    # inputstr = "KH YH"
    # names = getallnames("leader")
    # print(names)
    rawnames = ['康', '郑', '杨', '喻', '技术部', '工程部设备', '办公室', '办公室后勤', '办公室团委', '办公室XX']
    unknown,departnames,midnames,topnames = translatenames(rawnames)
    print(unknown,departnames,midnames,topnames)


if __name__ == "__main__":
    main()
