

# 汇总
# A X公司党委来文   8
# A X局党委来文    2
# B X公司团委来文   9
# B X局团委来文    3
# C X公司纪委来文   10
# C X局纪委来文    4
# D X公司行政来文   7
# D X局行政来文    1
# D 其他来文  12
# D Y分局来文    6
# E X公司工会来文   11
# E X局工会来文    5

# print("press A or a\t to choose folder 党委")
# print("press B or b\t to choose folder 团委")
# print("press C or c\t to choose folder 纪委")
# print("press D or d\t to choose folder 行政")
# print("press E or e\t to choose folder 工会")

def getcategory(recno):
    leadingnumbers=[[8,2],[9,3],[10,4],[7,1,12,6],[11,5]]
    categorys=["A","B","C","D","E"]
    mathingcategorys = [[categorys[i]]*len(leadingnumbers[i]) for i in range(len(leadingnumbers)) ]
    flatnumbers = [i for subl in leadingnumbers for i in subl]
    flatcategorys = [i for subl in mathingcategorys for i in subl]
    lookupdict = dict(zip(flatnumbers,flatcategorys))
    k = int(recno.split("-")[0])
    return lookupdict.get(k)


