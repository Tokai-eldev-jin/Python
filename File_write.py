#-------------------------------------------------------------------------------
# Name:        module1
# Purpose:
#
# Author:      A0044
#
# Created:     24/04/2023
# Copyright:   (c) A0044 2023
# Licence:     <your licence>
#-------------------------------------------------------------------------------


def main():
    pass





def File_write(status_num,file_id,持出者,確認者,koumoku_list,do_year,loginUser):

    import jin as j
    import global_value as g


    kazu:int
    i:int
    t:int
    no_flag:bool



    f1={}
    no_flag = True


    ##フォルダの作成
    if j.fopen(j.path_join(g.CUR_DIR,"user",koumoku_list[2]),"f_exists") == False:
        j.File_system("fs_createfolder", j.path_join(g.CUR_DIR,"user",koumoku_list[2]))


    if j.fopen(j.path_join(g.CUR_DIR,"user",koumoku_list[20]), "f_exists") == False:
        j.File_system("fs_createfolder", j.path_join(g.CUR_DIR,"user",koumoku_list[20]))


    if j.fopen(j.path_join(g.CUR_DIR,"user",koumoku_list[25] ), "f_exists") == False:
        j.File_system("fs_createfolder", j.path_join(g.CUR_DIR,"user",koumoku_list[25]))




    ##datafileに書き込み
    while True:
        #ダミーファイルがない時
        if j.fopen(j.path_join(g.CUR_DIR,"ini","dami_"+do_year+"_datafile.txt"), "f_exists") == False:
            #ダミーファイルの作成
            j.File_system("fs_createtextfile", j.path_join(g.CUR_DIR,"ini","dami_"+do_year+"_datafile.txt"))

            f1 = j.fopen(j.path_join(g.CUR_DIR,"ini",do_year+"_datafile.txt"), "f_read_or_f_write")
            kazu = j.fget(f1, -1)

            no_flag = True

            for i in range(1,kazu+1):
                getdata = j.fget(f1, i, 1) + "_" + j.fget(f1, i, 2)
                if getdata == file_id:
                    ##no_flag = False

                    ##status
                    j.fput(f1, status_num, i, 3 )

                    if status_num == 1 or status_num == 3:
                        ##data
                        for t in range(0,len(koumoku_list)):
                            j.fput(f1, j.hash_data_out(koumoku_list, "hash_val", t), i, 4 + t)

                    elif status_num == 4:
                        j.gettime(0)
                        j.fput(f1, g.G_TIME_YY4 + "/" + g.G_TIME_MM2 + "/" + g.G_TIME_DD2, i, 4 + 26)
                        j.fput(f1, g.G_TIME_HH2 + ":" + g.G_TIME_NN2, i, 4 + 27)

                    break

            j.fclose(f1)

            j.File_system("fs_deletefile", j.path_join(g.CUR_DIR,"ini","dami_" + do_year + "_datafile.txt"))

            break
        else:
            pass



    ##USER_fileに書き込み
    while True:
        if j.fopen(j.path_join(g.CUR_DIR,"user",持出者,"dami_" + 持出者 + ".txt"), "f_exists") == False:
            ##dami
            j.File_system("fs_createtextfile", j.path_join(g.CUR_DIR,"user",持出者,"dami_" + 持出者 + ".txt"))

            f1 = j.fopen(j.path_join(g.CUR_DIR,"user",持出者,持出者 + ".txt"), "f_read_or_f_write")
            kazu = j.fget(f1, -1)

            no_flag = True

            for i in range(1,kazu+1):
                getdata = j.fget(f1, i, 1) + "_" + j.fget(f1, i, 2)
                if getdata == file_id:
                    no_flag = False

                    ##status
                    j.fput(f1, status_num, i, 3)

                    if status_num == 1 or status_num == 3:
                        ##data
                        for t in range(0,len(koumoku_list)):
                            j.fput(f1, j.hash_data_out(koumoku_list, "hash_val", t), i, 4 + t)

                    elif status_num == 4:
                        j.gettime(0)
                        j.fput(f1, g.G_TIME_YY4 + "/" + g.G_TIME_MM2 + "/" + g.G_TIME_DD2, i, 4 + 26)
                        j.fput(f1, g.G_TIME_HH2 + ":" + g.G_TIME_NN2, i, 4 + 27)

                    break

            if no_flag == True:
                ##id
                g.token_str = file_id
                j.fput(f1, j.token("_", g.token_str), kazu + 1, 1)
                j.fput(f1, j.token("_", g.token_str), kazu + 1, 2)

                ##status
                j.fput(f1, status_num, kazu + 1, 3)

                ##data
                for t  in range(0 ,len(koumoku_list)):
                    j.fput(f1, j.hash_data_out(koumoku_list, "hash_val", t), kazu + 1, 4 + t)

            else:
                pass



            j.fclose(f1)


            j.File_system("fs_deletefile", j.path_join(g.CUR_DIR,"user" , 持出者 , "dami_" + 持出者 + ".txt"))

            break
        else:
            pass






    ##確認者_Re_fileに書き込み
    while True:
        if j.fopen(j.path_join(g.CUR_DIR,"user" , 確認者 , "dami_" + 確認者 + ".txt"), "f_exists") == False:
            ##dami
            j.File_system("fs_createtextfile", j.path_join(g.CUR_DIR,"user" , 確認者 , "dami_" + 確認者 + ".txt"))

            no_flag = True

            f1 = j.fopen(j.path_join(g.CUR_DIR,"user" , 確認者 , "Re_" + 確認者 + ".txt"), "f_read_or_f_write")
            kazu = j.fget(f1, -1)

            for i in range( 1 , kazu+1):
                getdata = j.fget(f1, i, 1) + "_" + j.fget(f1, i, 2)
                if getdata == file_id:

                    no_flag = False


                    if status_num == 1: j.fput(f1, 1, i, 3)           ##status 1にする
                    if status_num == 2: j.fput(f1, 0, i, 3)           ##tatus 0にする　（無効）
                    if status_num == 3: j.fput(f1, 3, i, 3)           ##status 3にする
                    if status_num == 4: j.fput(f1, 4, i, 3)           ##status 4にする
                    if status_num == 5: j.fput(f1, 0, i, 3)           ##status 0にする　（無効）
                    if status_num == 6: j.fput(f1, 3, i, 3)           ##status 3にする　（無効）
                    if status_num == 7: j.fput(f1, 0, i, 3)           ##status 0にする　（無効）

                    break



            if no_flag == True:
                ##id
                g.token_str = file_id
                j.fput(f1, j.token("_", g.token_str), kazu + 1, 1)
                j.fput(f1, j.token("_", g.token_str), kazu + 1, 2)

                if status_num == 1: j.fput(f1, 1, kazu + 1, 3)##status 1にする
                if status_num == 2: j.fput(f1, 0, kazu + 1, 3)##status 0にする　（無効）
                if status_num == 3: j.fput(f1, 3, kazu + 1, 3)##status 3にする
                if status_num == 4: j.fput(f1, 4, kazu + 1, 3)##status 4にする
                if status_num == 5: j.fput(f1, 0, kazu + 1, 3)##status 0にする　（無効）
                if status_num == 6: j.fput(f1, 3, kazu + 1, 3)##status 3にする　（無効）
                if status_num == 7: j.fput(f1, 0, kazu + 1, 3)##status 0にする　（無効）
            else:
                pass


            j.fclose(f1)

            j.File_system("fs_deletefile", j.path_join(g.CUR_DIR,"user" , 確認者 , "dami_" + 確認者 + ".txt"))

            break
        else:
            pass


    ##USER_Re_fileに書き込み
    while True:
        if j.fopen(j.path_join(g.CUR_DIR,"user" , 持出者 , "dami_" + 持出者 + ".txt"), "f_exists") == False:
            #dami
            j.File_system("fs_createtextfile", j.path_join(g.CUR_DIR,"user" , 持出者 , "dami_" + 持出者 + ".txt"))

            f1 = j.fopen(j.path_join(g.CUR_DIR,"user" , 持出者 , "Re_" + 持出者 + ".txt"), "f_read_or_f_write")
            kazu = j.fget(f1, -1)

            no_flag = True

            for i in range( 1 , kazu+1):
                getdata = j.fget(f1, i, 1) + "_" + j.fget(f1, i, 2)

                if getdata == file_id:

                    no_flag = False

                    if status_num == 1: j.fput(f1, 0, i, 3)##status 0にする  （無効）
                    if status_num == 2: j.fput(f1, 2, i, 3)##status 2にする
                    if status_num == 3: j.fput(f1, 0, i, 3)##status 0にする　（無効）
                    if status_num == 4: j.fput(f1, 4, i, 3)##status 4にする
                    if status_num == 5: j.fput(f1, 5, i, 3)##status 5にする　（無効）
                    if status_num == 6: j.fput(f1, 6, i, 3)##status 6にする　（無効）
                    if status_num == 7: j.fput(f1, 0, i, 3)##status 0にする　（無効）

                    break



            if no_flag == True:

                    getdata = file_id
                    g.token_str=getdata
                    j.fput(f1, j.token("_", g.token_str), kazu + 1, 1)
                    j.fput(f1, j.token("_", g.token_str), kazu + 1, 2)

                    if status_num == 1: j.fput(f1, 0, kazu + 1, 3)##status 0にする　（無効）
                    if status_num == 2: j.fput(f1, 2, kazu + 1, 3)##status 2にする
                    if status_num == 3: j.fput(f1, 0, kazu + 1, 3)##status 0にする　（無効）
                    if status_num == 4: j.fput(f1, 4, kazu + 1, 3)##status 4にする
                    if status_num == 5: j.fput(f1, 5, kazu + 1, 3)##status 5にする　（無効）
                    if status_num == 6: j.fput(f1, 6, kazu + 1, 3)##status 6にする　（無効）
                    if status_num == 7: j.fput(f1, 0, kazu + 1, 3)##status 0にする　（無効）
            else:
                pass



            j.fclose(f1)

            j.File_system("fs_deletefile", j.path_join(g.CUR_DIR,"user" , 持出者 , "dami_" + 持出者 + ".txt"))

            break
        else:
            pass





if __name__ == '__main__':
    main()
