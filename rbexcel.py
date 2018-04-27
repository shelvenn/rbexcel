# -*- coding: utf-8 -*-

import time
import pandas as pd
import matplotlib.pyplot as plt



alfa_1 = 0.2
alfa_2 = 1

MAX_COLUMN_LEN = 1048

new_df = pd.DataFrame()
res_df = pd.DataFrame()


def read_excel(xls_name, sheetname):
    # 代码示例:
    excel_path = xls_name
    d = pd.read_excel(excel_path, sheet_name=sheetname)
    print(d.keys())
    col_d1 = d['D1']

    MAX_COLUMN_LEN = len(col_d1)
    print('MAX_COLUMN_LEN:', MAX_COLUMN_LEN)

    return d, col_d1


def get_D_LLT_col(df):
    alfa_1 = 0.2

    d_idx = 0
    d_cols = df[0:]
    for di in range(len(d_cols.columns) - 1):
        d_idx = d_idx + 1
        col_name = 'D' + str(d_idx)
        d = d_cols[col_name]

        d_llt = [0.0] * MAX_COLUMN_LEN
        for t in range(len(d)):
            if t < 2:
                d_llt[t] = d[t]
            else:
                # if d[t-2] <= 0:
                #     d_llt[t] = d[t]
                # else:
                dd = (alfa_1 - alfa_1 * alfa_1 / 4) * d[t] + (alfa_1 * alfa_1 / 2) * d[t - 1]
                dd = dd - (alfa_1 - 3 * alfa_1 * alfa_1 / 4) * d[t - 2] + 2 * (1 - alfa_1) * d_llt[t - 1]
                d_llt[t] = dd - (1 - alfa_1) * (1 - alfa_1) * d_llt[t - 2]

        # 通过列表创建series
        llt_series_list = pd.Series(d_llt)
        llt_name = 'D-LLT-' + str(d_idx)
        new_df[llt_name] = llt_series_list


def get_g_col(df):
    alfa_1 = 0.2

    d_idx = 0
    d_cols = df[0:]
    for di in range(len(d_cols.columns) - 1):
        d_idx = d_idx + 1
        col_name = 'D' + str(d_idx)
        col_d = d_cols[col_name]

        g = [0.0] * MAX_COLUMN_LEN
        for t in range(len(col_d)):
            if t == 0:
                g[0] = 0.0
            elif col_d[t] > 0 and col_d[t - 1] > 0:
                g[t] = col_d[t] / col_d[t - 1] - 1
            else:
                g[t] = 0.0

        # 通过列表创建series
        g_series_list = pd.Series(g)
        g_name = 'G' + str(d_idx)
        new_df[g_name] = g_series_list

    G = [0.0] * MAX_COLUMN_LEN
    d_idx = 0
    new_d_cols = new_df[0:]
    for t in range(MAX_COLUMN_LEN - 1):
        d_idx = 0
        for di in range(len(d_cols.columns) - 1):
            d_idx = d_idx + 1
            g_name = 'G' + str(d_idx)
            G[t] = G[t] + new_df[g_name][t]

    G_series_col = pd.Series(G)
    new_df['G'] = G_series_col


#    return new_df

def get_n_col(df):
    d_idx = 0
    d_cols = df[0:]
    for di in range(len(d_cols.columns) - 1):
        d_idx = d_idx + 1
        col_name = 'D' + str(d_idx)
        d = d_cols[col_name]
        n_col = [0.0] * MAX_COLUMN_LEN
        for t in range(len(d)):
            if d[t] > 0:
                n_col[t] = 1.0
            else:
                n_col[t] = 0.0
        n_series_col = pd.Series(n_col)

        d_name = 'N' + str(d_idx)
        new_df[d_name] = n_series_col

    N = [0.0] * MAX_COLUMN_LEN
    d_idx = 0
    new_d_cols = new_df[0:]
    for t in range(MAX_COLUMN_LEN - 1):
        N[t] = 0.0
        for di in range(len(d_cols.columns) - 1):
            d_idx = d_idx + 1
            d_name = 'N' + str(d_idx)
            N[t] = N[t] + new_df[d_name][t]

        d_idx = 0

    N_series_col = pd.Series(N)
    new_df['N'] = N_series_col
    return new_df


def get_I_col(df):
    d_idx = 0
    _cols = df[0:]
    for di in range(len(_cols.columns) - 1):
        d_idx = d_idx + 1

        i_col = [0.0] * MAX_COLUMN_LEN
        new_cols = new_df[0:]
        #        col_name = 'G' + str(d_idx)
        g = new_cols['G']
        #        col_name = 'N' + str(d_idx)
        n = new_cols['N']
        for t in range(len(g) - 1):
            if t == 0:
                i_col[t] = 1
            else:
                gn = g[t] / n[t]
                i_col[t] = i_col[t - 1] * (1 + g[t] / n[t])

        i_series_col = pd.Series(i_col)

        new_df['I'] = i_series_col


def get_I_MA30_col(df):
    d_idx = 0
    _cols = df[0:]
    for di in range(len(_cols.columns) - 1):
        d_idx = d_idx + 1

        ima30_col = [0.0] * MAX_COLUMN_LEN
        new_cols = new_df[0:]
        i_col = new_cols['I']
        for t in range(len(i_col) - 1):
            if t < 29:
                ima30_col[t] = i_col[t]
            else:
                i_list = i_col.tolist()[t - 29:t + 1]
                new_isum = sum(i_list)
                ima30_col[t] = new_isum / 30
                # print('t:', t, '   len:',len(i_list),'    ima30_col:',ima30_col[t],'   new_isum:',new_isum)

        ima30_series_col = pd.Series(ima30_col)
        new_df['I-MA30'] = ima30_series_col


def get_I_MA40_col(df):
    d_idx = 0
    _cols = df[0:]
    for di in range(len(_cols.columns) - 1):
        d_idx = d_idx + 1

        ima40_col = [0.0] * MAX_COLUMN_LEN
        new_cols = new_df[0:]
        i_col = new_cols['I']
        for t in range(len(i_col) - 1):
            if t < 39:
                ima40_col[t] = i_col[t]
            else:
                i_list = i_col.tolist()[t - 39:t + 1]
                new_isum = sum(i_list)
                ima40_col[t] = new_isum / 40

        ima40_series_col = pd.Series(ima40_col)
        new_df['I-MA40'] = ima40_series_col


def get_I_MA60_col(df):
    d_idx = 0
    _cols = df[0:]
    for di in range(len(_cols.columns) - 1):
        d_idx = d_idx + 1

        ima60_col = [0.0] * MAX_COLUMN_LEN
        new_cols = new_df[0:]
        i_col = new_cols['I']
        for t in range(len(i_col) - 1):
            if t < 59:
                ima60_col[t] = i_col[t]
            else:
                i_list = i_col.tolist()[t - 59:t + 1]
                new_isum = sum(i_list)
                ima60_col[t] = new_isum / 60

        ima60_series_col = pd.Series(ima60_col)
        new_df['I-MA60'] = ima60_series_col


def get_I_LLT_col(df):
    illt_col = [0.0] * MAX_COLUMN_LEN
    new_cols = new_df[0:]
    i = new_cols['I']
    for t in range(len(i) - 1):
        if t < 3:
            illt_col[t] = i[t]
        else:
            illt = (alfa_1 - alfa_1 * alfa_1 / 4) * i[t] + (alfa_1 * alfa_1 / 2) * i[t - 1]
            illt = illt - (alfa_1 - 3 * alfa_1 * alfa_1 / 4) * i[t - 2]
            illt_col[t] = illt + 2 * (1 - alfa_1) * illt_col[t - 1] - (1 - alfa_1) * (1 - alfa_1) * illt_col[t - 2]

    illt_series_col = pd.Series(illt_col)
    new_df['I-LLT'] = illt_series_col


def get_Fst_cn_col(df):
    fst_cn_col = [0.0] * MAX_COLUMN_LEN
    new_cols = new_df[0:]  # 获取df所有的列
    ima30 = new_cols['I-MA30']
    for t in range(len(ima30) - 1):
        if t == 0:
            fst_cn_col[t] = 0.0
        else:
            if ima30[t] >= ima30[t - 1]:
                fst_cn_col[t] = 1.0
            else:
                fst_cn_col[t] = 0.0

    fst_cn_series_col = pd.Series(fst_cn_col)
    new_df['Fst-cn'] = fst_cn_series_col


def get_Fst_CN_col(df):
    d_idx = 0
    d_cols = df[0:]
    new_cols = new_df[0:]
    for di in range(len(d_cols.columns) - 1):
        d_idx = d_idx + 1
        col_name = 'D' + str(d_idx)
        d = d_cols[col_name]
        fst_CN_col = [0.0] * MAX_COLUMN_LEN
        col_name = 'G' + str(d_idx)
        g = new_cols[col_name]
        for t in range(len(d)):
            if t == 0:
                fst_CN_col[t] = 0.0
            elif d[t] == d[t - 1] or abs(g[t]) > 0.09:
                fst_CN_col[t] = fst_CN_col[t - 1]
            else:
                fst_CN_col[t] = new_cols['Fst-cn'][t]
        fst_CN_series_col = pd.Series(fst_CN_col)

        d_name = 'Fst-CN-' + str(d_idx)
        new_df[d_name] = fst_CN_series_col


def get_Fst_SL_col(df):
    d_idx = 0
    d_cols = df[0:]
    new_cols = new_df[0:]
    for di in range(len(d_cols.columns) - 1):
        d_idx = d_idx + 1
        col_name = 'D' + str(d_idx)
        d = d_cols[col_name]
        fst_SL_col = [0.0] * MAX_COLUMN_LEN
        col_name = 'Fst-CN-' + str(d_idx)
        fst_CN_col = new_cols[col_name]
        col_name = 'G' + str(d_idx)
        g = new_cols[col_name]
        for t in range(len(d)):
            if t < 2:
                fst_SL_col[t] = d[t]
            else:
                if d[t - 1] <= 0:
                    fst_SL_col[t] = d[t]
                else:
                    fst_SL_col[t] = fst_SL_col[t - 1] * (1 + fst_CN_col[t - 2] * g[t])

        fst_SL_series_col = pd.Series(fst_SL_col)
        d_name = 'Fst-SL-' + str(d_idx)
        new_df[d_name] = fst_SL_series_col


def get_Fst_MA20_col(df):
    d_idx = 0
    d_cols = df[0:]
    for di in range(len(d_cols.columns) - 1):
        d_idx = d_idx + 1
        col_name = 'D' + str(d_idx)
        d = d_cols[col_name]

        fstma20_col = [0.0] * MAX_COLUMN_LEN
        new_cols = new_df[0:]  # 获取df所有的列
        i = new_cols['I']
        col_name = 'Fst-SL-' + str(d_idx)
        fst_SL_col = new_cols[col_name]
        for t in range(len(i) - 1):
            if t < 19:
                fstma20_col[t] = fst_SL_col[t]
            else:
                if d[t - 19] <= 0:
                    fstma20_col[t] = fst_SL_col[t]
                else:
                    # isum = 0.0
                    # for it in range(t + 1):
                    #     if it >= t - 19:
                    #         isum = isum + fst_SL_col[it]
                    # fstma20_col[t] = isum / 20
                    i_list = fst_SL_col.tolist()[t - 19:t + 1]
                    new_isum = sum(i_list)
                    fstma20_col[t] = new_isum / 60

        fstma20_series_col = pd.Series(fstma20_col)
        col_name = 'Fst-MA20-' + str(d_idx)
        new_df[col_name] = fstma20_series_col


def get_Snd_cn_col(df):
    d_idx = 0
    d_cols = df[0:]
    new_cols = new_df[0:]
    for di in range(len(d_cols.columns) - 1):
        d_idx = d_idx + 1
        col_name = 'D' + str(d_idx)
        d = d_cols[col_name]
        snd_cn_col = [0.0] * MAX_COLUMN_LEN
        col_name = 'Fst-MA20-' + str(d_idx)
        fst_MA20_col = new_cols[col_name]
        for t in range(len(d)):
            if t == 0:
                snd_cn_col[t] = 0
            else:
                if fst_MA20_col[t] < fst_MA20_col[t - 1]:
                    snd_cn_col[t] = 0
                else:
                    snd_cn_col[t] = 1

        snd_cn_series_col = pd.Series(snd_cn_col)
        d_name = 'Snd-cn-' + str(d_idx)
        new_df[d_name] = snd_cn_series_col


def get_Snd_CN_col(df):
    d_idx = 0
    d_cols = df[0:]
    new_cols = new_df[0:]
    for di in range(len(d_cols.columns) - 1):
        d_idx = d_idx + 1
        col_name = 'D' + str(d_idx)
        d = d_cols[col_name]
        snd_CN_col = [0.0] * MAX_COLUMN_LEN
        col_name = 'Snd-cn-' + str(d_idx)
        snd_cn_col = new_cols[col_name]
        col_name = 'G' + str(d_idx)
        g = new_cols[col_name]
        for t in range(len(d)):
            if t == 0:
                snd_CN_col[t] = 0
            else:
                if d[t] == d[t - 1] or abs(g[t]) > 0.09:
                    snd_CN_col[t] = snd_CN_col[t - 1]
                else:
                    snd_CN_col[t] = snd_cn_col[t]

        snd_CN_series_col = pd.Series(snd_CN_col)
        d_name = 'Snd-CN-' + str(d_idx)
        new_df[d_name] = snd_CN_series_col


def get_Snd_SL_col(df):
    d_idx = 0
    d_cols = df[0:]
    new_cols = new_df[0:]
    for di in range(len(d_cols.columns) - 1):
        d_idx = d_idx + 1
        col_name = 'D' + str(d_idx)
        d = d_cols[col_name]
        snd_SL_col = [0.0] * MAX_COLUMN_LEN
        col_name = 'Fst-SL-' + str(d_idx)
        fst_sl_col = new_cols[col_name]
        col_name = 'Snd-CN-' + str(d_idx)
        snd_CN_col = new_cols[col_name]
        col_name = 'G' + str(d_idx)
        g = new_cols[col_name]
        for t in range(len(d)):
            if t < 2:
                snd_SL_col[t] = fst_sl_col[t]
            else:
                if d[t - 1] <= 0:
                    snd_SL_col[t] = fst_sl_col[t]
                else:
                    snd_SL_col[t] = snd_SL_col[t - 1] * (1 + snd_CN_col[t - 2] * g[t])

        snd_SL_series_col = pd.Series(snd_SL_col)
        d_name = 'Snd-SL-' + str(d_idx)
        new_df[d_name] = snd_SL_series_col


def get_Snd_LLT_col(df):
    d_idx = 0
    d_cols = df[0:]
    new_cols = new_df[0:]
    for di in range(len(d_cols.columns) - 1):
        d_idx = d_idx + 1
        col_name = 'D' + str(d_idx)
        d = d_cols[col_name]
        snd_LLT_col = [0.0] * MAX_COLUMN_LEN
        col_name = 'Snd-SL-' + str(d_idx)
        snd_SL_col = new_cols[col_name]
        col_name = 'Snd-CN-' + str(d_idx)
        snd_CN_col = new_cols[col_name]
        col_name = 'G' + str(d_idx)
        g = new_cols[col_name]
        for t in range(len(d)):
            if t < 2:
                snd_LLT_col[t] = snd_SL_col[t]
            else:
                # if snd_SL_col[t-2] <= 0:
                #     snd_LLT_col[t] = snd_SL_col[t]
                # else:
                llt = (alfa_1 - alfa_1 * alfa_1 / 4) * snd_SL_col[t] + (alfa_1 * alfa_1 / 2) * snd_SL_col[t - 1]
                llt = llt - (alfa_1 - 3 * alfa_1 * alfa_1 / 4) * snd_SL_col[t - 2]
                snd_LLT_col[t] = llt + 2 * (1 - alfa_1) * snd_LLT_col[t - 1] - (1 - alfa_1) * (1 - alfa_1) * snd_LLT_col[t - 2]


        snd_LLT_series_col = pd.Series(snd_LLT_col)
        d_name = 'Snd-LLT-' + str(d_idx)
        new_df[d_name] = snd_LLT_series_col


def get_VS_CN_col(df):
    alfa_2 = 1
    d_idx = 0
    d_cols = df[0:]
    new_cols = new_df[0:]
    for di in range(len(d_cols.columns) - 1):
        d_idx = d_idx + 1
        col_name = 'D' + str(d_idx)
        d = d_cols[col_name]
        vs_CN_col = [0.0] * MAX_COLUMN_LEN
        col_name = 'D-LLT-' + str(d_idx)
        d_llt_col = new_cols[col_name]
        col_name = 'I-LLT'
        i_llt_col = new_cols[col_name]
        for t in range(len(d)):
            if t < 4:
                vs_CN_col[t] = 1
            else:
                if (d_llt_col[t - 1] == 0) or (d_llt_col[t - 4] == 0):
                    vs_CN_col[t] = 0
                elif (d_llt_col[t] / d_llt_col[t - 1] - 1 > alfa_2 * (i_llt_col[t] / i_llt_col[t - 1] - 1)) and (
                        d_llt_col[t] / d_llt_col[t - 4] - 1 > alfa_2 * (i_llt_col[t] / i_llt_col[t - 4] - 1)):
                    vs_CN_col[t] = 1
                else:
                    vs_CN_col[t] = 0

        vs_CN_series_col = pd.Series(vs_CN_col)
        d_name = 'VS-CN-' + str(d_idx)
        new_df[d_name] = vs_CN_series_col


def get_COMP_CN_col(df):
    d_idx = 0
    d_cols = df[0:]
    new_cols = new_df[0:]
    for di in range(len(d_cols.columns) - 1):
        d_idx = d_idx + 1
        col_name = 'D' + str(d_idx)
        d = d_cols[col_name]
        comp_CN_col = [0.0] * MAX_COLUMN_LEN
        col_name = 'VS-CN-' + str(d_idx)
        vs_cn_col = new_cols[col_name]
        col_name = 'G' + str(d_idx)
        g = new_cols[col_name]
        for t in range(len(d)):
            if t < 1:
                comp_CN_col[t] = 0
            else:
                if d[t] == d[t - 1] or abs(g[t]) > 0.09:
                    comp_CN_col[t] = comp_CN_col[t - 1]
                else:
                    comp_CN_col[t] = vs_cn_col[t]

        comp_CN_series_col = pd.Series(comp_CN_col)
        d_name = 'COMP-CN-' + str(d_idx)
        new_df[d_name] = comp_CN_series_col


def get_M_W_col(df):
    d_idx = 0
    d_cols = df[0:]
    new_cols = new_df[0:]

    M_W_col = [0.0] * MAX_COLUMN_LEN
    col_name = 'I-MA40'
    i_ma40_col = new_cols[col_name]
    col_name = 'I-MA60'
    i_ma60_col = new_cols[col_name]

    for t in range(len(i_ma40_col)):
        if i_ma40_col[t] < i_ma60_col[t]:
            M_W_col[t] = -1
        else:
            M_W_col[t] = 1

    M_W_series_col = pd.Series(M_W_col)
    d_name = 'M-W'
    new_df[d_name] = M_W_series_col


def get_W_col(df):
    d_idx = 0
    d_cols = df[0:]
    new_cols = new_df[0:]
    # for di in range(len(d_cols.columns)-1):
    #     d_idx = d_idx + 1
    #     col_name = 'D' + str(d_idx)
    #     d = d_cols[col_name]
    W_col = [0.0] * MAX_COLUMN_LEN
    col_name = 'M-W'
    m_w_col = new_cols[col_name]
    for t in range(len(m_w_col)):
        if t == 0:
            W_col[t] = 0
        else:
            d_idx = 0
            ccsum = 0
            for di in range(len(d_cols.columns) - 1):
                d_idx = d_idx + 1
                col_name = 'COMP-CN-' + str(d_idx)
                comp_cn_col = new_cols[col_name]
                ccsum = ccsum + comp_cn_col[t]

            if m_w_col[t] > 0 and ccsum > 0:
                W_col[t] = 1 / ccsum
                if W_col[t] > 0.1:
                    W_col[t] = 0.1
            else:
                W_col[t] = 0

    W_series_col = pd.Series(W_col)
    d_name = 'W'
    new_df[d_name] = W_series_col


'''
if  Snd_CN(t,j) > 0
         n(t,j) = 1
         N(t,1) = sum(n(t,j)		 
EQUAASS(t,1) = EQUAASS(t-1,1)  +  EQUAASS(t-1,1) * SUM(Snd_CN(t-2,j)*g(t,j))/N(t-2,1)
'''


def get_EQUAASS_col(df):  ## 这个EQUAASS_col是我加的，麻烦文勰帮我检查一下代码有没有问题
    d_idx = 0
    d_cols = df[0:]
    new_cols = new_df[0:]
    equaass_col = [0.0] * MAX_COLUMN_LEN
    col_name = 'I'
    i_col = new_cols[col_name]
    col_name = 'W'
    w_col = new_cols[col_name]
    col_name = 'N'
    n_col = new_cols[col_name]
    for t in range(len(i_col)):
        if t < 2:
            equaass_col[t] = i_col[t]
        else:
            d_idx = 0
            ccsum = 0.0
            cnsum = 0.0
            for di in range(len(d_cols.columns) - 1):
                d_idx = d_idx + 1
                col_name = 'Snd-CN-' + str(d_idx)
                snd_CN_col = new_cols[col_name]
                col_name = 'G' + str(d_idx)
                g_col = new_cols[col_name]
                ccsum = ccsum + snd_CN_col[t - 2] * g_col[t]
                # 计算 sum(snd_CN_col[t-2])
                cnsum = cnsum + snd_CN_col[t - 2]
            if cnsum == 0:
                equaass_col[t] = equaass_col[t - 1] + equaass_col[t - 1] * ccsum
            else:
                equaass_col[t] = equaass_col[t - 1] + equaass_col[t - 1] * (1 / cnsum) * ccsum
    equaass_series_col = pd.Series(equaass_col)
    d_name = 'EQUAASS'
    new_df[d_name] = equaass_series_col
    res_df[d_name] = equaass_series_col  ##我加的一段到此结束


'''
这一段是原来的代码，只是将原来的ASS命名为COMPASS，其他没动，因涉及到有一个EQUAASS,其他因素的命名会不会冲突?
'''


def get_COMPASS_col(df):  # 这一段为了区别于上一段，在ASS前加了COMP,不会有问题吧？
    d_idx = 0
    d_cols = df[0:]
    new_cols = new_df[0:]
    compass_col = [0.0] * MAX_COLUMN_LEN
    col_name = 'I'
    i_col = new_cols[col_name]
    col_name = 'W'
    w_col = new_cols[col_name]
    col_name = 'N'
    n_col = new_cols[col_name]
    for t in range(len(i_col)):
        if t < 2:
            compass_col[t] = i_col[t]
        else:
            if w_col[t - 2] == 0:
                d_idx = 0
                ccsum = 0.0
                for di in range(len(d_cols.columns) - 1):
                    d_idx = d_idx + 1
                    col_name = 'Snd-CN-' + str(d_idx)
                    snd_CN_col = new_cols[col_name]
                    col_name = 'G' + str(d_idx)
                    g_col = new_cols[col_name]
                    ccsum = ccsum + snd_CN_col[t - 2] * g_col[t]
                compass_col[t] = compass_col[t - 1] + compass_col[t - 1] * (1 / n_col[t]) * ccsum
            else:
                d_idx = 0
                ccsum = 0.0
                for di in range(len(d_cols.columns) - 1):
                    d_idx = d_idx + 1
                    col_name = 'COMP-CN-' + str(d_idx)
                    comp_CN_col = new_cols[col_name]
                    col_name = 'G' + str(d_idx)
                    g_col = new_cols[col_name]
                    ccsum = ccsum + comp_CN_col[t - 2] * g_col[t]
                compass_col[t] = compass_col[t - 1] + compass_col[t - 1] * w_col[t - 2] * ccsum

    compass_series_col = pd.Series(compass_col)
    d_name = 'COMPASS'
    new_df[d_name] = compass_series_col
    res_df[d_name] = compass_series_col


'''
第一步:计算EQUAASS的MA5
for  t<5
     EQUAASS_MA5(t) = EQUAASS(t)
	t>=5
	IF EQUAASS_MA5(t-4) > 0
	 EQUAASS_MA5(t) = sum( EQUAASS(t:t-4)/5
	else   
	  EQUAASS_MA5(t) = EQUAASS(t)
'''
def get_EQUAASS_MA5_col(df):  # 这一段为了区别于上一段，在ASS前加了COMP,不会有问题吧？
    d_idx = 0
    d_cols = df[0:]
    new_cols = new_df[0:]
    EQUAASS_MA5_col = [0.0] * MAX_COLUMN_LEN
    col_name = 'EQUAASS'
    EQUAASS_col = new_cols[col_name]
    for t in range(len(EQUAASS_col)):
        if t < 5:
            EQUAASS_MA5_col[t] = EQUAASS_col[t]
        else:
            if EQUAASS_MA5_col[t - 4] > 0:
                EQUAASS_MA5_col[t] = 0
            else:
                EQUAASS_MA5_col[t] = EQUAASS_col[t]

    EQUAASS_MA5_series_col = pd.Series(EQUAASS_MA5_col)
    d_name = 'EQUAASS_MA5'
    new_df[d_name] = EQUAASS_MA5_series_col


'''
第二步：计算EQUAASS的LLT
    alfa = 0.1                  
	for t in range(len(i) - 1):
        if t < 3:
            EQUAASSLLT_col[t] = EQUAASS[t]
        else:
            EQUAASSllt = (alfa_1-alfa_1*alfa_1/4)*EQUAASS[t] + (alfa_1*alfa_1/2)*EQUAASS[t-1]
            EQUAASSllt = EQUAASSllt - (alfa_1-3*alfa_1*alfa_1/4)*EQUAASS[t-2]
            EQUAASSllt_col[t] = EQUAASSllt + 2*(1-alfa_1)*EQUAASSllt_col[t-1] - (1-alfa_1)*(1-alfa_1)*EQUAASSllt_col[t-2]	
'''


def get_EQUAASS_LLT_col(df):
    d_idx = 0
    d_cols = df[0:]
    alfa_llt = 0.1
    new_cols = new_df[0:]
    EQUAASS_LLT_col = [0.0] * MAX_COLUMN_LEN
    col_name = 'EQUAASS'
    EQUAASS_col = new_cols[col_name]
    for t in range(len(EQUAASS_col)):
        if t < 3:
            EQUAASS_LLT_col[t] = EQUAASS_col[t]
        else:
            EQUAASSllt = (alfa_llt - alfa_llt * alfa_llt / 4) * EQUAASS_col[t] + (alfa_llt * alfa_llt / 2) * \
                         EQUAASS_col[t - 1]
            EQUAASSllt = EQUAASSllt - (alfa_llt - 3 * alfa_llt * alfa_llt / 4) * EQUAASS_col[t - 2]
            EQUAASS_LLT_col[t] = EQUAASSllt + 2 * (1 - alfa_llt) * EQUAASS_LLT_col[t - 1] - (1 - alfa_llt) * (
                        1 - alfa_llt) * EQUAASS_LLT_col[t - 2]

    EQUAASS_LLT_series_col = pd.Series(EQUAASS_LLT_col)
    d_name = 'EQUAASS_LLT'
    new_df[d_name] = EQUAASS_LLT_series_col


'''
第三步：计算EQUACOR
    for  t < 5
		EQUACOR(t) = 1
	for  t >= 5	
        if EQUAASS_MA5(t-1) > EQUAASSllt_col(t-1) or EQUAASSllt_col(t) > EQUAASSllt_col(t-1):
            EQUACOR(t) = 1
        else
            EQUACOR(t) = 0	
'''
# def get_EQUACOR_col(df):                                    ## 添加一个EQUACOR
#     d_idx = 0
#     d_cols = df[0:]
#     alfa_llt = 0.1
#     new_cols = new_df[0:]
#     EQUACOR = [0.0] * MAX_COLUMN_LEN
#     col_name = 'EQUAASS_MA5'
#     EQUAASS_MA5_col = new_cols[col_name]
#     col_name = 'EQUAASS_LLT'
#     EQUAASS_LLT_col = new_cols[col_name]
#     for t in range(len(EQUAASS_MA5_col)):
#         if t < 5:
#             EQUACOR[t] = 1.0
#         else:
#             if EQUAASS_MA5_col[t - 1] > EQUAASS_LLT_col[t - 1] or EQUAASS_LLT_col[t] > EQUAASS_LLT_col[t - 1]:
#                 EQUACOR[t] = 1.0
#             else:
#                 EQUACOR[t] = 0.0
#
#     EQUACOR_series_col = pd.Series(EQUACOR)
#     d_name = 'EQUACOR'
#     new_df[d_name] = EQUACOR_series_col
#     res_df[d_name] = EQUACOR_series_col
'''
这里稼接了COMPCOR的部分计算公式，也是可用的，就怕语法错误，麻烦文勰帮我看看有没有语法错误吧	
'''


def get_EQUACOR_col(df):
    d_idx = 0
    d_cols = df[0:]
    alfa_llt = 0.1
    new_cols = new_df[0:]
    EQUACOR_col = [0] * MAX_COLUMN_LEN
    col_name = 'EQUAASS_MA5'
    EQUAASS_MA5_col = new_cols[col_name]
    col_name = 'EQUAASS_LLT'
    EQUAASS_LLT_col = new_cols[col_name]
    col_name = 'EQUAASS'
    EQUAASS_col = new_cols[col_name]
    for t in range(len(EQUAASS_MA5_col)):
        if t >= 11:
            # equaass_max = EQUAASS_col[t]
            # for tt in range(t - 10, t):
            #     equaass_max = max(equaass_max, EQUAASS_col[tt])
            tmp_list = EQUAASS_col.tolist()[t-10:t+1]
            equaass_max = max(tmp_list)
            if EQUAASS_col[t] > equaass_max * 0.95:
                EQUACOR_col[t] = 1
        if t >= 11:
            # equaass_max = EQUAASS_col[t]
            # equaass_min = EQUAASS_col[t]
            # for tt in range(t - 10, t):
            #     equaass_max = max(equaass_max, EQUAASS_col[tt])
            #     equaass_min = min(equaass_min, EQUAASS_col[tt])
            tmp_list = EQUAASS_col.tolist()[t-10:t+1]
            equaass_max = max(tmp_list)
            equaass_min = max(tmp_list)
            if (EQUAASS_col[t] < equaass_max * 0.9) and (EQUAASS_col[t] > equaass_min * 1.01):
                EQUACOR_col[t] = 1

    EQUACOR_series_col = pd.Series(EQUACOR_col)
    d_name = 'EQUACOR'
    new_df[d_name] = EQUACOR_series_col
    res_df[d_name] = EQUACOR_series_col


## 加入ASS的LLT(alfa = 0.1)；加入ASS的MA5  COR1的条件是ASS的MA5t-1 > ASS 的 LLTt-1，或者，ASS的LLTt > ASS 的LLTt-1,
#    COR1为 1 ，否则，为 0


def get_F_EQUAASS_col(df):
    d_idx = 0
    d_cols = df[0:]
    new_cols = new_df[0:]
    f_equaass_col = [0.0] * MAX_COLUMN_LEN
    col_name = 'I'
    i_col = new_cols[col_name]
    col_name = 'W'
    w_col = new_cols[col_name]
    col_name = 'N'
    n_col = new_cols[col_name]
    col_name = 'EQUACOR'
    EQUACOR_col = new_cols[col_name]

    for t in range(len(i_col)):
        f_equaass_col[t] = i_col[t]

    for t in range(len(i_col)):
        if t >= 2:
            d_idx = 0
            ccsum = 0.0
            cnsum = 0.0
            for di in range(len(d_cols.columns) - 1):
                d_idx = d_idx + 1
                col_name = 'Snd-CN-' + str(d_idx)
                snd_CN_col = new_cols[col_name]
                col_name = 'G' + str(d_idx)
                g_col = new_cols[col_name]
                ccsum = ccsum + snd_CN_col[t - 2] * g_col[t]
                cnsum = cnsum + snd_CN_col[t - 2]
            if cnsum == 0:
                f_equaass_col[t] = f_equaass_col[t - 1] + f_equaass_col[t - 1] * ccsum * EQUACOR_col[t - 2]
            else:
                f_equaass_col[t] = f_equaass_col[t - 1] + f_equaass_col[t - 1] * ccsum * EQUACOR_col[t - 2] / cnsum

    f_equaass_series_col = pd.Series(f_equaass_col)
    d_name = 'F-EQUAASS'
    new_df[d_name] = f_equaass_series_col
    res_df[d_name] = f_equaass_series_col



def get_COMPCOR_col(df):
    d_idx = 0
    d_cols = df[0:]
    new_cols = new_df[0:]

    COMPCOR_col = [0.0] * MAX_COLUMN_LEN
    col_name = 'I'
    i_col = new_cols[col_name]
    col_name = 'COMPASS'
    compass_col = new_cols[col_name]

    for t in range(len(i_col)):
        if t < 11:
            COMPCOR_col[t] = 1.0
        else:
            COMPCOR_col[t] = 0.0

    for t in range(len(i_col)):
        if t >= 11:
            # i_max = i_col[t]
            # for tt in range(t - 10, t):
            #     i_max = max(i_max, i_col[tt])
            tmp_list = i_col.tolist()[t-10:t+1]
            i_max = max(tmp_list)
            if i_col[t] > i_max * 0.95:
                COMPCOR_col[t] = 1.0
        if t >= 10:
            # i_max = i_col[t]
            # i_min = i_col[t]
            # for tt in range(t - 10, t):
            #     i_max = max(i_max, i_col[tt])
            #     i_min = min(i_min, i_col[tt])

            tmp_list = i_col.tolist()[t-10:t+1]
            i_max = max(tmp_list)
            i_min = max(tmp_list)
            if (i_col[t] < i_max * 0.9) and (i_col[t] > i_min * 1.01):
                COMPCOR_col[t] = 1.0
        if t >= 11:
            # compass_max = compass_col[t]
            # for tt in range(t - 10, t):
            #     compass_max = max(compass_max, compass_col[tt])

            tmp_list = compass_col.tolist()[t-10:t+1]
            compass_max = max(tmp_list)
            if compass_col[t] > compass_max * 0.95:
                COMPCOR_col[t] = 1.0
        if t >= 59:
            # compass_max = compass_col[t]
            # for tt in range(t - 50, t - 1):
            #     compass_max = max(compass_max, compass_col[tt])
            tmp_list = compass_col.tolist()[t-50:t+1]
            compass_max = max(tmp_list)
            if compass_col[t] > compass_max * 0.95:
                COMPCOR_col[t] = 1.0
        if t >= 11:
            # compass_max = compass_col[t]
            # compass_min = compass_col[t]
            # for tt in range(t - 10, t):
            #     compass_max = max(compass_max, compass_col[tt])
            #     compass_min = min(compass_min, compass_col[tt])
            tmp_list = compass_col.tolist()[t-10:t+1]
            compass_max = max(tmp_list)
            compass_min = max(tmp_list)
            if (compass_col[t] < compass_max * 0.9) and (compass_col[t] > compass_min * 1.01):
                COMPCOR_col[t] = 1.0
        if t >= 59:
            # compass_max = compass_col[t]
            # compass_min = compass_col[t]
            # for tt in range(t - 50, t - 1):
            #     compass_max = max(compass_max, compass_col[tt])
            #     compass_min = min(compass_min, compass_col[tt])
            tmp_list = compass_col.tolist()[t-50:t+1]
            compass_max = max(tmp_list)
            compass_min = max(tmp_list)
            if (compass_col[t] < compass_max * 0.9) and (compass_col[t] > compass_min * 1.01):
                COMPCOR_col[t] = 1.0
        if t >= 99:
            # compass_max = compass_col[t]
            # compass_min = compass_col[t]
            # for tt in range(t - 90, t - 1):
            #     compass_max = max(compass_max, compass_col[tt])
            #     compass_min = min(compass_min, compass_col[tt])
            tmp_list = compass_col.tolist()[t-90:t+1]
            compass_max = max(tmp_list)
            compass_min = max(tmp_list)
            if (compass_col[t] < compass_max * 0.9) and (compass_col[t] > compass_min * 1.01):
                COMPCOR_col[t] = 1.0

    COMPCOR_series_col = pd.Series(COMPCOR_col)
    d_name = 'COMPCOR'
    new_df[d_name] = COMPCOR_series_col
    res_df[d_name] = COMPCOR_series_col


def get_F_COMPASS_col(df):
    d_idx = 0
    d_cols = df[0:]
    new_cols = new_df[0:]
    f_compass_col = [0.0] * MAX_COLUMN_LEN
    col_name = 'I'
    i_col = new_cols[col_name]
    col_name = 'W'
    w_col = new_cols[col_name]
    col_name = 'N'
    n_col = new_cols[col_name]
    col_name = 'COMPCOR'
    COMPCOR_col = new_cols[col_name]

    for t in range(len(i_col)):
        f_compass_col[t] = i_col[t]

    for t in range(len(i_col)):
        if t >= 2:
            if w_col[t - 2] == 0:
                d_idx = 0
                ccsum = 0.0
                for di in range(len(d_cols.columns) - 1):
                    d_idx = d_idx + 1
                    col_name = 'Snd-CN-' + str(d_idx)
                    snd_CN_col = new_cols[col_name]
                    col_name = 'G' + str(d_idx)
                    g_col = new_cols[col_name]
                    ccsum = ccsum + snd_CN_col[t - 2] * g_col[t]
                f_compass_col[t] = f_compass_col[t - 1] + f_compass_col[t - 1] * ccsum * COMPCOR_col[t - 2] / n_col[t]
            else:
                d_idx = 0
                ccsum = 0.0
                for di in range(len(d_cols.columns) - 1):
                    d_idx = d_idx + 1
                    col_name = 'COMP-CN-' + str(d_idx)
                    comp_CN_col = new_cols[col_name]
                    col_name = 'G' + str(d_idx)
                    g_col = new_cols[col_name]
                    ccsum = ccsum + comp_CN_col[t - 2] * g_col[t]
                f_compass_col[t] = f_compass_col[t - 1] + f_compass_col[t - 1] * w_col[t - 2] * COMPCOR_col[t - 2] * ccsum

    f_compass_series_col = pd.Series(f_compass_col)
    d_name = 'F-COMPASS'
    new_df[d_name] = f_compass_series_col
    res_df[d_name] = f_compass_series_col


def get_EQUACN_col(df):  # 这里加了一个EQUACN，麻烦文勰帮我看看有没有问题
    d_idx = 0
    d_cols = df[0:]
    new_cols = new_df[0:]
    for di in range(len(d_cols.columns) - 1):
        d_idx = d_idx + 1
        col_name = 'D' + str(d_idx)
        d = d_cols[col_name]
        EQUACN_col = [0.0] * MAX_COLUMN_LEN
        col_name = 'Snd-CN-' + str(d_idx)
        snd_CN_col = new_cols[col_name]

        col_name = 'W'
        w_col = new_cols[col_name]
        col_name = 'N'
        n_col = new_cols[col_name]
        col_name = 'EQUACOR'
        equacor_col = new_cols[col_name]
        for t in range(len(d)):
            if t < 2:
                EQUACN_col[t] = 0.0
            else:
                EQUACN_col[t] = equacor_col[t] * snd_CN_col[t] / n_col[t]

        EQUACN_series_col = pd.Series(EQUACN_col)
        d_name = 'EQUACN-' + str(d_idx)
        new_df[d_name] = EQUACN_series_col
        res_df[d_name] = EQUACN_series_col  # 增加的EQUACN到此结束


'''
EQUACN(t,j) = equacor(t,1)*Snd_CN(t,j)/n(t,1)
'''


def get_COMPCN_col(df):  # 这里修改为COMPCN
    d_idx = 0
    d_cols = df[0:]
    new_cols = new_df[0:]
    for di in range(len(d_cols.columns) - 1):
        d_idx = d_idx + 1
        col_name = 'D' + str(d_idx)
        d = d_cols[col_name]
        COMPCN_col = [0.0] * MAX_COLUMN_LEN
        col_name = 'Snd-CN-' + str(d_idx)
        snd_CN_col = new_cols[col_name]
        col_name = 'COMP-CN-' + str(d_idx)
        comp_CN_col = new_cols[col_name]
        col_name = 'W'
        w_col = new_cols[col_name]
        col_name = 'N'
        n_col = new_cols[col_name]
        col_name = 'COMPCOR'
        compcor_col = new_cols[col_name]
        for t in range(len(d)):
            if t < 2:
                COMPCN_col[t] = 0.0
            else:
                if w_col[t] == 0:
                    COMPCN_col[t] = compcor_col[t] * snd_CN_col[t] / n_col[t]
                elif w_col[t] > 0:
                    COMPCN_col[t] = w_col[t] * compcor_col[t] * comp_CN_col[t]

        COMPCN_series_col = pd.Series(COMPCN_col)
        d_name = 'COMPCN-' + str(d_idx)
        new_df[d_name] = COMPCN_series_col
        res_df[d_name] = COMPCN_series_col


if __name__ == '__main__':

    begin_time = time.time()

    df, col_d1 = read_excel(r'D:\test\data3.xlsx', 'Sheet1')

    get_D_LLT_col(df)
    print('get_D_LLT_col')
    get_g_col(df)
    print('get_g_col')
    get_n_col(df)
    print('get_n_col')
    #    print(new_df['N1'], new_df['N2'], new_df['N3'], new_df['N4'])

    get_I_col(df)
    print('get_I_col')
    get_I_MA30_col(df)
    print('get_I_MA30_col')
    get_I_MA40_col(df)
    print('get_I_MA40_col')
    get_I_MA60_col(df)
    print('get_I_MA60_col')
    get_I_LLT_col(df)
    print('get_I_LLT_col')
    get_Fst_cn_col(df)
    print('get_Fst_cn_col')
    get_Fst_CN_col(df)
    print('get_Fst_CN_col')
    get_Fst_SL_col(df)
    print('get_Fst_SL_col')
    get_Fst_MA20_col(df)
    print('get_Fst_MA20_col')
    get_Snd_cn_col(df)
    print('get_Snd_cn_col')
    get_Snd_CN_col(df)
    print('get_Snd_CN_col')
    get_Snd_SL_col(df)
    print('get_Snd_SL_col')
    get_Snd_LLT_col(df)
    print('get_Snd_LLT_col')
    get_VS_CN_col(df)
    print('get_VS_CN_col')
    get_COMP_CN_col(df)
    print('get_COMP_CN_col')
    get_M_W_col(df)
    print('get_M_W_col')
    get_W_col(df)
    print('get_W_col')
    get_EQUAASS_col(df)
    print('get_EQUAASS_col')
    get_COMPASS_col(df)
    print('get_COMPASS_col')
    get_EQUAASS_MA5_col(df)
    print('get_EQUAASS_MA5_col')
    get_EQUAASS_LLT_col(df)
    print('get_EQUAASS_LLT_col')
    get_COMPCOR_col(df)
    print('get_COMPCOR_col')
    get_F_COMPASS_col(df)
    print('get_F_COMPASS_col')
    get_EQUACOR_col(df)
    print('get_EQUACOR_col')
    get_F_EQUAASS_col(df)
    print('get_F_EQUAASS_col')
    get_EQUACN_col(df)
    print('get_EQUACN_col')
    get_COMPCN_col(df)
    print('get_COMPCN_col')

    end_time = time.time()
    num_of_sec = end_time - begin_time
    print('运行时间: %d分%d秒  共%d秒' % (num_of_sec / 60, num_of_sec % 60, num_of_sec))

    xlsx_path = 'F:/python/'
    # new_df.to_excel(xlsx_path + 'new_df.xlsx', sheet_name='new_sheet')
    # res_df.to_excel(xlsx_path + 'res_df.xlsx', sheet_name='sheet1')



    #plt.rcParams['font.sas-serig'] = ['SimHei']
    plt.figure(1)
    plt.subplot(221)
    plt.xlabel('t')
    plt.ylabel('F-EQUAASS')
    plt.plot(res_df.index, res_df['F-EQUAASS'])
    plt.title('F-EQUAASS')

    plt.subplot(222)
    plt.xlabel('t')
    plt.ylabel('I')
    plt.plot(new_df.index, new_df['I'])
    plt.title('I')

    plt.subplot(223)
    plt.xlabel('t')
    plt.ylabel('F-EQUAASS')
    plt.plot(res_df.index, res_df['F-EQUAASS'])
    plt.title('F-EQUAASS')

    plt.subplot(224)
    plt.xlabel('t')
    plt.ylabel('F-COMPASS')
    plt.plot(new_df.index, new_df['F-COMPASS'])
    plt.title('F-COMPASS')

    plt.tight_layout()
    #plt.plot(new_df.index, new_df['I'])
    plt.savefig(xlsx_path + "result_pic.jpg")
    plt.show()


