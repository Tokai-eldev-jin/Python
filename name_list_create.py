def main():
    print("")




def name_list_create(DB_path1):

    import jin as j
    import global_value as g


    f1={}

    f1=j.fopen(DB_path1,"f_read")
    endline=j.fget(f1,-1)

    DBm=j.Hairetu(endline,3)

    gyou=0
    for i in range(1,endline+1):
        DBm[gyou][0]=j.fget(f1,i,3)     #ひらがな
        DBm[gyou][1]=j.fget(f1,i,1)     #従業員番号
        DBm[gyou][2]=j.fget(f1,i,2)     #名前
        gyou=gyou+1


    DBm = j.Hairetu_sort(DBm,"reverse")
    return DBm


if __name__ == '__main__':
    main()

