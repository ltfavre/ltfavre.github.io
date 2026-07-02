

# ============================================================
#same as v6, tried to fix the issue that the sample items are not correcly parsed


import json
import pandas as pd
from pathlib import Path
import re
import traceback
from datetime import datetime

# ============================================================
# === UPDATE THESE TWO PATHS ONLY                           ===
# ============================================================
input_path = Path(
    r"C:\Users\UFAVRLA\PythonProjects\KTEA_CELF_LLM_Scoring_Lab\data\raw\fake_wisc_protocol_extracted.json"
)

output_excel_path = Path(
    r"C:\Users\UFAVRLA\PythonProjects\KTEA_CELF_LLM_Scoring_Lab\outputs\mainv12_14_5.xlsx"
)

# ============================================================
# === FIELDS TO EXTRACT                                     ===
# ============================================================
fields_to_extract = [
    "mr_st", "mr_et", "id", "mr_r", "mr_l",
    "sr_st", "sr_et", "sr_p19", "sr_p20",
    "fw_l", "fw_st", "fw_et", "fw_r",
    "in_p12", "in_st", "in_p13", "in_et",

    # === NEW FIELDS 1 ===
    "si_p1", "si_p2", "si_st", "si_et",
    "vc_st", "vc_et", "vc_p7", "vc_p8",
    "co_st", "co_et", "co_p17", "co_p18",
# === NEW FIELDS 2 ===
    "bd_st", "bd_et", "df_st", "df_et","df",
    "ds_st", "ds_et", "cd_st", "cd_et",
    "vp_st", "vp_et", "rd_st", "rd_et",
    "ss_st", "ss_et", "ar_st", "ar_et",
    "db_st", "db_et", "ssp_st", "ssp_et",
    "nsq_st", "nsq_et",
    "sa_st", "sa_et", "lns_st", "lns_et",
    "mdf_1", "mdf_2", "mdf_3", "mdf_4","mdf_5", "mdf_6", "mdf_7", "mdf_text","gn",
    "vc_p7_1-3","bd_p3_4", "bd_p3_5-8", "bd_p3_9", "bd_p3_10-14", "lns_scores",
    "sa_1_5", "sa_6_14", "sa_15_23", "nsq", "db_p15_scores", "db_p14_scores", "df_scores", "ds_scores",
    "bd_p2", "vp_r", "vp_l", "rd", "ar_r", "db_p14", "db_p15", "ssp_p16", "ssp_p15", "lns", "ar_l", "ds",
    "cd_t", "ss_t", "nsq_t1", "nsq_t2"
]
# === FIELDS THAT MUST ALWAYS BE *_raw                      ===
# ============================================================
ALWAYS_RAW_FIELDS = {
    "mr_st", "mr_et",
    "fw_st", "fw_et",
    "in_st", "in_et",
    "sr_st", "sr_et",

    # === NEW timing fields that should always be stored as *_raw ===
    "si_st", "si_et",
    "vc_st", "vc_et",
    "co_st", "co_et",

# === NEW timing fields that should always be stored as *_raw  2===
    "bd_st", "bd_et", "df_st", "df_et",
    "ds_st", "ds_et", "cd_st", "cd_et",
    "vp_st", "vp_et", "rd_st", "rd_et",
    "ss_st", "ss_et", "ar_st", "ar_et",
    "db_st", "db_et", "ssp_st", "ssp_et",
    "nsq_st", "nsq_et", "ssp_st", "ssp_et",
    "sa_st", "sa_et", "lns_st", "lns_et"
}

# ============================================================
# === FINAL OUTPUT COLUMN ORDER (AUTHORITATIVE)             ===
# ============================================================
# Paste your full desired column order here, one per line.
# This is the order that will be enforced in the final Excel output.
FINAL_COLUMN_ORDER_TEXT = r"""
id
si_st_raw
si_p1_Sample_verbatim
si_p1_Sample_dk
si_p1_Sample_nr
si_p1_Sample_s0
si_p1_Sample_s1
si_p1_Sample_s2
si_p1_1_verbatim
si_p1_1_dk
si_p1_1_nr
si_p1_1_s0
si_p1_1_s1
si_p1_1_s2
si_p1_2_verbatim
si_p1_2_dk
si_p1_2_nr
si_p1_2_s0
si_p1_2_s1
si_p1_2_s2
si_p1_3_verbatim
si_p1_3_dk
si_p1_3_nr
si_p1_3_s0
si_p1_3_s1
si_p1_3_s2
si_p1_4_verbatim
si_p1_4_dk
si_p1_4_nr
si_p1_4_s0
si_p1_4_s1
si_p1_4_s2
si_p1_5_verbatim
si_p1_5_dk
si_p1_5_nr
si_p1_5_s0
si_p1_5_s1
si_p1_5_s2
si_p1_6_verbatim
si_p1_6_dk
si_p1_6_nr
si_p1_6_s0
si_p1_6_s1
si_p1_6_s2
si_p1_7_verbatim
si_p1_7_dk
si_p1_7_nr
si_p1_7_s0
si_p1_7_s1
si_p1_7_s2
si_p1_8_verbatim
si_p1_8_dk
si_p1_8_nr
si_p1_8_s0
si_p1_8_s1
si_p1_8_s2
si_p1_9_verbatim
si_p1_9_dk
si_p1_9_nr
si_p1_9_s0
si_p1_9_s1
si_p1_9_s2
si_p1_10_verbatim
si_p1_10_dk
si_p1_10_nr
si_p1_10_s0
si_p1_10_s1
si_p1_10_s2
si_p1_11_verbatim
si_p1_11_dk
si_p1_11_nr
si_p1_11_s0
si_p1_11_s1
si_p1_11_s2
si_p2_12_verbatim
si_p2_12_dk
si_p2_12_nr
si_p2_12_s0
si_p2_12_s1
si_p2_12_s2
si_p2_13_verbatim
si_p2_13_dk
si_p2_13_nr
si_p2_13_s0
si_p2_13_s1
si_p2_13_s2
si_p2_14_verbatim
si_p2_14_dk
si_p2_14_nr
si_p2_14_s0
si_p2_14_s1
si_p2_14_s2
si_p2_15_verbatim
si_p2_15_dk
si_p2_15_nr
si_p2_15_s0
si_p2_15_s1
si_p2_15_s2
si_p2_16_verbatim
si_p2_16_dk
si_p2_16_nr
si_p2_16_s0
si_p2_16_s1
si_p2_16_s2
si_p2_17_verbatim
si_p2_17_dk
si_p2_17_nr
si_p2_17_s0
si_p2_17_s1
si_p2_17_s2
si_p2_18_verbatim
si_p2_18_dk
si_p2_18_nr
si_p2_18_s0
si_p2_18_s1
si_p2_18_s2
si_et_raw
bd_st_raw
bd_p2_1_time_1
bd_p2_1_time_2
bd_p2_1_de
bd_p2_1_nbc_0
bd_p2_1_nbc_1
bd_p2_1_nbc_2
bd_p2_1_nbc_3
bd_p2_1_nbc_4
bd_p2_1_re
bd_p2_1_s0
bd_p2_1_s1
bd_p2_1_s2
bd_p2_2_time_1
bd_p2_2_time_2
bd_p2_2_de
bd_p2_2_nbc_0
bd_p2_2_nbc_1
bd_p2_2_nbc_2
bd_p2_2_nbc_3
bd_p2_2_nbc_4
bd_p2_2_re
bd_p2_2_s0
bd_p2_2_s1
bd_p2_2_s2
bd_p2_3_time_1
bd_p2_3_time_2
bd_p2_3_de
bd_p2_3_nbc_0
bd_p2_3_nbc_1
bd_p2_3_nbc_2
bd_p2_3_nbc_3
bd_p2_3_nbc_4
bd_p2_3_re
bd_p2_3_s0
bd_p2_3_s1
bd_p2_3_s2
bd_p3_4_4_time_1
bd_p3_4_4_time_2
bd_p3_4_4_de
bd_p3_4_4_nbc_0
bd_p3_4_4_nbc_1
bd_p3_4_4_nbc_2
bd_p3_4_4_nbc_3
bd_p3_4_4_nbc_4
bd_p3_4_4_re
bd_p3_4_4_s0
bd_p3_4_4_s1
bd_p3_4_4_s2
bd_p3_5-8_5_time
bd_p3_5-8_5_de
bd_p3_5-8_5_nbc_0
bd_p3_5-8_5_nbc_1
bd_p3_5-8_5_nbc_2
bd_p3_5-8_5_nbc_3
bd_p3_5-8_5_nbc_4
bd_p3_5-8_5_re
bd_p3_5-8_5_s0
bd_p3_5-8_5_s4
bd_p3_5-8_6_time
bd_p3_5-8_6_de
bd_p3_5-8_6_nbc_0
bd_p3_5-8_6_nbc_1
bd_p3_5-8_6_nbc_2
bd_p3_5-8_6_nbc_3
bd_p3_5-8_6_nbc_4
bd_p3_5-8_6_re
bd_p3_5-8_6_s0
bd_p3_5-8_6_s4
bd_p3_5-8_7_time
bd_p3_5-8_7_de
bd_p3_5-8_7_nbc_0
bd_p3_5-8_7_nbc_1
bd_p3_5-8_7_nbc_2
bd_p3_5-8_7_nbc_3
bd_p3_5-8_7_nbc_4
bd_p3_5-8_7_re
bd_p3_5-8_7_s0
bd_p3_5-8_7_s4
bd_p3_5-8_8_time
bd_p3_5-8_8_de
bd_p3_5-8_8_nbc_0
bd_p3_5-8_8_nbc_1
bd_p3_5-8_8_nbc_2
bd_p3_5-8_8_nbc_3
bd_p3_5-8_8_nbc_4
bd_p3_5-8_8_re
bd_p3_5-8_8_s0
bd_p3_5-8_8_s4
bd_p3_9_9_time
bd_p3_9_9_de
bd_p3_9_9_nbc_0
bd_p3_9_9_nbc_1
bd_p3_9_9_nbc_2
bd_p3_9_9_nbc_3
bd_p3_9_9_nbc_4
bd_p3_9_9_re
bd_p3_9_9_s0
bd_p3_9_9_s4
bd_p3_9_9_s5
bd_p3_9_9_s6
bd_p3_9_9_s7
bd_p3_10-14_10_time
bd_p3_10-14_10_de
bd_p3_10-14_10_nbc_0
bd_p3_10-14_10_nbc_1
bd_p3_10-14_10_nbc_2
bd_p3_10-14_10_nbc_3
bd_p3_10-14_10_nbc_4
bd_p3_10-14_10_nbc_5
bd_p3_10-14_10_nbc_6
bd_p3_10-14_10_nbc_7
bd_p3_10-14_10_nbc_8
bd_p3_10-14_10_nbc_9
bd_p3_10-14_10_re
bd_p3_10-14_10_s0
bd_p3_10-14_10_s4
bd_p3_10-14_10_s5
bd_p3_10-14_10_s6
bd_p3_10-14_10_s7
bd_p3_10-14_11_time
bd_p3_10-14_11_de
bd_p3_10-14_11_nbc_0
bd_p3_10-14_11_nbc_1
bd_p3_10-14_11_nbc_2
bd_p3_10-14_11_nbc_3
bd_p3_10-14_11_nbc_4
bd_p3_10-14_11_nbc_5
bd_p3_10-14_11_nbc_6
bd_p3_10-14_11_nbc_7
bd_p3_10-14_11_nbc_8
bd_p3_10-14_11_nbc_9
bd_p3_10-14_11_re
bd_p3_10-14_11_s0
bd_p3_10-14_11_s4
bd_p3_10-14_11_s5
bd_p3_10-14_11_s6
bd_p3_10-14_11_s7
bd_p3_10-14_12_time
bd_p3_10-14_12_de
bd_p3_10-14_12_nbc_0
bd_p3_10-14_12_nbc_1
bd_p3_10-14_12_nbc_2
bd_p3_10-14_12_nbc_3
bd_p3_10-14_12_nbc_4
bd_p3_10-14_12_nbc_5
bd_p3_10-14_12_nbc_6
bd_p3_10-14_12_nbc_7
bd_p3_10-14_12_nbc_8
bd_p3_10-14_12_nbc_9
bd_p3_10-14_12_re
bd_p3_10-14_12_s0
bd_p3_10-14_12_s4
bd_p3_10-14_12_s5
bd_p3_10-14_12_s6
bd_p3_10-14_12_s7
bd_p3_10-14_13_time
bd_p3_10-14_13_de
bd_p3_10-14_13_nbc_0
bd_p3_10-14_13_nbc_1
bd_p3_10-14_13_nbc_2
bd_p3_10-14_13_nbc_3
bd_p3_10-14_13_nbc_4
bd_p3_10-14_13_nbc_5
bd_p3_10-14_13_nbc_6
bd_p3_10-14_13_nbc_7
bd_p3_10-14_13_nbc_8
bd_p3_10-14_13_nbc_9
bd_p3_10-14_13_re
bd_p3_10-14_13_s0
bd_p3_10-14_13_s4
bd_p3_10-14_13_s5
bd_p3_10-14_13_s6
bd_p3_10-14_13_s7
bd_p3_10-14_14_time
bd_p3_10-14_14_de
bd_p3_10-14_14_nbc_0
bd_p3_10-14_14_nbc_1
bd_p3_10-14_14_nbc_2
bd_p3_10-14_14_nbc_3
bd_p3_10-14_14_nbc_4
bd_p3_10-14_14_nbc_5
bd_p3_10-14_14_nbc_6
bd_p3_10-14_14_nbc_7
bd_p3_10-14_14_nbc_8
bd_p3_10-14_14_nbc_9
bd_p3_10-14_14_re
bd_p3_10-14_14_s0
bd_p3_10-14_14_s4
bd_p3_10-14_14_s5
bd_p3_10-14_14_s6
bd_p3_10-14_14_s7
bd_et_raw
mr_st_raw
mr_l_SA_r1
mr_l_SA_r2
mr_l_SA_r3
mr_l_SA_r4
mr_l_SA_r5
mr_l_SA_dk
mr_l_SA_nr
mr_l_SA_s0
mr_l_SA_s1
mr_l_SB_r1
mr_l_SB_r2
mr_l_SB_r3
mr_l_SB_r4
mr_l_SB_r5
mr_l_SB_dk
mr_l_SB_nr
mr_l_SB_s0
mr_l_SB_s1
mr_l_1_r1
mr_l_1_r2
mr_l_1_r3
mr_l_1_r4
mr_l_1_r5
mr_l_1_dk
mr_l_1_nr
mr_l_1_s0
mr_l_1_s1
mr_l_2_r1
mr_l_2_r2
mr_l_2_r3
mr_l_2_r4
mr_l_2_r5
mr_l_2_dk
mr_l_2_nr
mr_l_2_s0
mr_l_2_s1
mr_l_3_r1
mr_l_3_r2
mr_l_3_r3
mr_l_3_r4
mr_l_3_r5
mr_l_3_dk
mr_l_3_nr
mr_l_3_s0
mr_l_3_s1
mr_l_4_r1
mr_l_4_r2
mr_l_4_r3
mr_l_4_r4
mr_l_4_r5
mr_l_4_dk
mr_l_4_nr
mr_l_4_s0
mr_l_4_s1
mr_l_5_r1
mr_l_5_r2
mr_l_5_r3
mr_l_5_r4
mr_l_5_r5
mr_l_5_dk
mr_l_5_nr
mr_l_5_s0
mr_l_5_s1
mr_l_6_r1
mr_l_6_r2
mr_l_6_r3
mr_l_6_r4
mr_l_6_r5
mr_l_6_dk
mr_l_6_nr
mr_l_6_s0
mr_l_6_s1
mr_l_7_r1
mr_l_7_r2
mr_l_7_r3
mr_l_7_r4
mr_l_7_r5
mr_l_7_dk
mr_l_7_nr
mr_l_7_s0
mr_l_7_s1
mr_l_8_r1
mr_l_8_r2
mr_l_8_r3
mr_l_8_r4
mr_l_8_r5
mr_l_8_dk
mr_l_8_nr
mr_l_8_s0
mr_l_8_s1
mr_l_9_r1
mr_l_9_r2
mr_l_9_r3
mr_l_9_r4
mr_l_9_r5
mr_l_9_dk
mr_l_9_nr
mr_l_9_s0
mr_l_9_s1
mr_l_10_r1
mr_l_10_r2
mr_l_10_r3
mr_l_10_r4
mr_l_10_r5
mr_l_10_dk
mr_l_10_nr
mr_l_10_s0
mr_l_10_s1
mr_l_11_r1
mr_l_11_r2
mr_l_11_r3
mr_l_11_r4
mr_l_11_r5
mr_l_11_dk
mr_l_11_nr
mr_l_11_s0
mr_l_11_s1
mr_l_12_r1
mr_l_12_r2
mr_l_12_r3
mr_l_12_r4
mr_l_12_r5
mr_l_12_dk
mr_l_12_nr
mr_l_12_s0
mr_l_12_s1
mr_l_13_r1
mr_l_13_r2
mr_l_13_r3
mr_l_13_r4
mr_l_13_r5
mr_l_13_dk
mr_l_13_nr
mr_l_13_s0
mr_l_13_s1
mr_l_14_r1
mr_l_14_r2
mr_l_14_r3
mr_l_14_r4
mr_l_14_r5
mr_l_14_dk
mr_l_14_nr
mr_l_14_s0
mr_l_14_s1
mr_r_15_r1
mr_r_15_r2
mr_r_15_r3
mr_r_15_r4
mr_r_15_r5
mr_r_15_dk
mr_r_15_nr
mr_r_15_s0
mr_r_15_s1
mr_r_16_r1
mr_r_16_r2
mr_r_16_r3
mr_r_16_r4
mr_r_16_r5
mr_r_16_dk
mr_r_16_nr
mr_r_16_s0
mr_r_16_s1
mr_r_17_r1
mr_r_17_r2
mr_r_17_r3
mr_r_17_r4
mr_r_17_r5
mr_r_17_dk
mr_r_17_nr
mr_r_17_s0
mr_r_17_s1
mr_r_18_r1
mr_r_18_r2
mr_r_18_r3
mr_r_18_r4
mr_r_18_r5
mr_r_18_dk
mr_r_18_nr
mr_r_18_s0
mr_r_18_s1
mr_r_19_r1
mr_r_19_r2
mr_r_19_r3
mr_r_19_r4
mr_r_19_r5
mr_r_19_dk
mr_r_19_nr
mr_r_19_s0
mr_r_19_s1
mr_r_20_r1
mr_r_20_r2
mr_r_20_r3
mr_r_20_r4
mr_r_20_r5
mr_r_20_dk
mr_r_20_nr
mr_r_20_s0
mr_r_20_s1
mr_r_21_r1
mr_r_21_r2
mr_r_21_r3
mr_r_21_r4
mr_r_21_r5
mr_r_21_dk
mr_r_21_nr
mr_r_21_s0
mr_r_21_s1
mr_r_22_r1
mr_r_22_r2
mr_r_22_r3
mr_r_22_r4
mr_r_22_r5
mr_r_22_dk
mr_r_22_nr
mr_r_22_s0
mr_r_22_s1
mr_r_23_r1
mr_r_23_r2
mr_r_23_r3
mr_r_23_r4
mr_r_23_r5
mr_r_23_dk
mr_r_23_nr
mr_r_23_s0
mr_r_23_s1
mr_r_24_r1
mr_r_24_r2
mr_r_24_r3
mr_r_24_r4
mr_r_24_r5
mr_r_24_dk
mr_r_24_nr
mr_r_24_s0
mr_r_24_s1
mr_r_25_r1
mr_r_25_r2
mr_r_25_r3
mr_r_25_r4
mr_r_25_r5
mr_r_25_dk
mr_r_25_nr
mr_r_25_s0
mr_r_25_s1
mr_r_26_r1
mr_r_26_r2
mr_r_26_r3
mr_r_26_r4
mr_r_26_r5
mr_r_26_dk
mr_r_26_nr
mr_r_26_s0
mr_r_26_s1
mr_et_raw
df_st_raw
df_1_1_verbatim
df_1_1_dk
df_1_1_nr
df_1_1_t0
df_1_1_t1_2
df_1_2_verbatim
df_1_2_dk
df_1_2_nr
df_1_2_t0
df_1_2_t1_2
df_2_1_verbatim
df_2_1_dk
df_2_1_nr
df_2_1_t0
df_2_1_t1_2
df_2_2_verbatim
df_2_2_dk
df_2_2_nr
df_2_2_t0
df_2_2_t1_2
df_3_1_verbatim
df_3_1_dk
df_3_1_nr
df_3_1_t0
df_3_1_t1_2
df_3_2_verbatim
df_3_2_dk
df_3_2_nr
df_3_2_t0
df_3_2_t1_2
df_4_1_verbatim
df_4_1_dk
df_4_1_nr
df_4_1_t0
df_4_1_t1_2
df_4_2_verbatim
df_4_2_dk
df_4_2_nr
df_4_2_t0
df_4_2_t1_2
df_5_1_verbatim
df_5_1_dk
df_5_1_nr
df_5_1_t0
df_5_1_t1_2
df_5_2_verbatim
df_5_2_dk
df_5_2_nr
df_5_2_t0
df_5_2_t1_2
df_6_1_verbatim
df_6_1_dk
df_6_1_nr
df_6_1_t0
df_6_1_t1_2
df_6_2_verbatim
df_6_2_dk
df_6_2_nr
df_6_2_t0
df_6_2_t1_2
df_7_1_verbatim
df_7_1_dk
df_7_1_nr
df_7_1_t0
df_7_1_t1_2
df_7_2_verbatim
df_7_2_dk
df_7_2_nr
df_7_2_t0
df_7_2_t1_2
df_8_1_verbatim
df_8_1_dk
df_8_1_nr
df_8_1_t0
df_8_1_t1_2
df_8_2_verbatim
df_8_2_dk
df_8_2_nr
df_8_2_t0
df_8_2_t1_2
df_9_1_verbatim
df_9_1_dk
df_9_1_nr
df_9_1_t0
df_9_1_t1_2
df_9_2_verbatim
df_9_2_dk
df_9_2_nr
df_9_2_t0
df_9_2_t1_2
df_10_1_verbatim
df_10_1_dk
df_10_1_nr
df_10_1_t0
df_10_1_t1_2
df_10_2_verbatim
df_10_2_dk
df_10_2_nr
df_10_2_t0
df_10_2_t1_2
df_scores_1_s0
df_scores_1_s1_2
df_scores_1_s2_4
df_scores_2_s0
df_scores_2_s1_2
df_scores_2_s2_4
df_scores_3_s0
df_scores_3_s1_2
df_scores_3_s2_4
df_scores_4_s0
df_scores_4_s1_2
df_scores_4_s2_4
df_scores_5_s0
df_scores_5_s1_2
df_scores_5_s2_4
df_scores_6_s0
df_scores_6_s1_2
df_scores_6_s2_4
df_scores_7_s0
df_scores_7_s1_2
df_scores_7_s2_4
df_scores_8_s0
df_scores_8_s1_2
df_scores_8_s2_4
df_scores_9_s0
df_scores_9_s1_2
df_scores_9_s2_4
df_scores_10_s0
df_scores_10_s1_2
df_scores_10_s2_4
df_et_raw
ds_st_raw
ds_sa_1_verbatim
ds_sa_1_dk
ds_sa_1_nr
ds_sa_1_t0
ds_sa_1_t1_2
ds_sa_2_verbatim
ds_sa_2_dk
ds_sa_2_nr
ds_sa_2_t0
ds_sa_2_t1_2
ds_1_1_verbatim
ds_1_1_dk
ds_1_1_nr
ds_1_1_t0
ds_1_1_t1_2
ds_1_2_verbatim
ds_1_2_dk
ds_1_2_nr
ds_1_2_t0
ds_1_2_t1_2
ds_2_1_verbatim
ds_2_1_dk
ds_2_1_nr
ds_2_1_t0
ds_2_1_t1_2
ds_2_2_verbatim
ds_2_2_dk
ds_2_2_nr
ds_2_2_t0
ds_2_2_t1_2
ds_3_1_verbatim
ds_3_1_dk
ds_3_1_nr
ds_3_1_t0
ds_3_1_t1_2
ds_3_2_verbatim
ds_3_2_dk
ds_3_2_nr
ds_3_2_t0
ds_3_2_t1_2
ds_4_1_verbatim
ds_4_1_dk
ds_4_1_nr
ds_4_1_t0
ds_4_1_t1_2
ds_4_2_verbatim
ds_4_2_dk
ds_4_2_nr
ds_4_2_t0
ds_4_2_t1_2
ds_5_1_verbatim
ds_5_1_dk
ds_5_1_nr
ds_5_1_t0
ds_5_1_t1_2
ds_5_2_verbatim
ds_5_2_dk
ds_5_2_nr
ds_5_2_t0
ds_5_2_t1_2
ds_6_1_verbatim
ds_6_1_dk
ds_6_1_nr
ds_6_1_t0
ds_6_1_t1_2
ds_6_2_verbatim
ds_6_2_dk
ds_6_2_nr
ds_6_2_t0
ds_6_2_t1_2
ds_7_1_verbatim
ds_7_1_dk
ds_7_1_nr
ds_7_1_t0
ds_7_1_t1_2
ds_7_2_verbatim
ds_7_2_dk
ds_7_2_nr
ds_7_2_t0
ds_7_2_t1_2
ds_8_1_verbatim
ds_8_1_dk
ds_8_1_nr
ds_8_1_t0
ds_8_1_t1_2
ds_8_2_verbatim
ds_8_2_dk
ds_8_2_nr
ds_8_2_t0
ds_8_2_t1_2
ds_9_1_verbatim
ds_9_1_dk
ds_9_1_nr
ds_9_1_t0
ds_9_1_t1_2
ds_9_2_verbatim
ds_9_2_dk
ds_9_2_nr
ds_9_2_t0
ds_9_2_t1_2
ds_10_1_verbatim
ds_10_1_dk
ds_10_1_nr
ds_10_1_t0
ds_10_1_t1_2
ds_10_2_verbatim
ds_10_2_dk
ds_10_2_nr
ds_10_2_t0
ds_10_2_t1_2
ds_scores_1_s0
ds_scores_1_s1_2
ds_scores_1_s2_4
ds_scores_2_s0
ds_scores_2_s1_2
ds_scores_2_s2_4
ds_scores_3_s0
ds_scores_3_s1_2
ds_scores_3_s2_4
ds_scores_4_s0
ds_scores_4_s1_2
ds_scores_4_s2_4
ds_scores_5_s0
ds_scores_5_s1_2
ds_scores_5_s2_4
ds_scores_6_s0
ds_scores_6_s1_2
ds_scores_6_s2_4
ds_scores_7_s0
ds_scores_7_s1_2
ds_scores_7_s2_4
ds_scores_8_s0
ds_scores_8_s1_2
ds_scores_8_s2_4
ds_scores_9_s0
ds_scores_9_s1_2
ds_scores_9_s2_4
ds_scores_10_s0
ds_scores_10_s1_2
ds_scores_10_s2_4
ds_et_raw
cd_st_raw
cd_t
cd_et_raw
vc_st_raw
vc_p7_1-3_1_verbatim
vc_p7_1-3_1_dk
vc_p7_1-3_1_nr
vc_p7_1-3_1_s0
vc_p7_1-3_1_s1
vc_p7_1-3_2_verbatim
vc_p7_1-3_2_dk
vc_p7_1-3_2_nr
vc_p7_1-3_2_s0
vc_p7_1-3_2_s1
vc_p7_1-3_3_verbatim
vc_p7_1-3_3_dk
vc_p7_1-3_3_nr
vc_p7_1-3_3_s0
vc_p7_1-3_3_s1
vc_p7_4_verbatim
vc_p7_4_dk
vc_p7_4_nr
vc_p7_4_s0
vc_p7_4_s1
vc_p7_4_s2
vc_p7_5_verbatim
vc_p7_5_dk
vc_p7_5_nr
vc_p7_5_s0
vc_p7_5_s1
vc_p7_5_s2
vc_p7_6_verbatim
vc_p7_6_dk
vc_p7_6_nr
vc_p7_6_s0
vc_p7_6_s1
vc_p7_6_s2
vc_p7_7_verbatim
vc_p7_7_dk
vc_p7_7_nr
vc_p7_7_s0
vc_p7_7_s1
vc_p7_7_s2
vc_p7_8_verbatim
vc_p7_8_dk
vc_p7_8_nr
vc_p7_8_s0
vc_p7_8_s1
vc_p7_8_s2
vc_p7_9_verbatim
vc_p7_9_dk
vc_p7_9_nr
vc_p7_9_s0
vc_p7_9_s1
vc_p7_9_s2
vc_p7_10_verbatim
vc_p7_10_dk
vc_p7_10_nr
vc_p7_10_s0
vc_p7_10_s1
vc_p7_10_s2
vc_p7_11_verbatim
vc_p7_11_dk
vc_p7_11_nr
vc_p7_11_s0
vc_p7_11_s1
vc_p7_11_s2
vc_p7_12_verbatim
vc_p7_12_dk
vc_p7_12_nr
vc_p7_12_s0
vc_p7_12_s1
vc_p7_12_s2
vc_p8_13_verbatim
vc_p8_13_dk
vc_p8_13_nr
vc_p8_13_s0
vc_p8_13_s1
vc_p8_13_s2
vc_p8_14_verbatim
vc_p8_14_dk
vc_p8_14_nr
vc_p8_14_s0
vc_p8_14_s1
vc_p8_14_s2
vc_p8_15_verbatim
vc_p8_15_dk
vc_p8_15_nr
vc_p8_15_s0
vc_p8_15_s1
vc_p8_15_s2
vc_p8_16_verbatim
vc_p8_16_dk
vc_p8_16_nr
vc_p8_16_s0
vc_p8_16_s1
vc_p8_16_s2
vc_p8_17_verbatim
vc_p8_17_dk
vc_p8_17_nr
vc_p8_17_s0
vc_p8_17_s1
vc_p8_17_s2
vc_p8_18_verbatim
vc_p8_18_dk
vc_p8_18_nr
vc_p8_18_s0
vc_p8_18_s1
vc_p8_18_s2
vc_p8_19_verbatim
vc_p8_19_dk
vc_p8_19_nr
vc_p8_19_s0
vc_p8_19_s1
vc_p8_19_s2
vc_p8_20_verbatim
vc_p8_20_dk
vc_p8_20_nr
vc_p8_20_s0
vc_p8_20_s1
vc_p8_20_s2
vc_p8_21_verbatim
vc_p8_21_dk
vc_p8_21_nr
vc_p8_21_s0
vc_p8_21_s1
vc_p8_21_s2
vc_p8_22_verbatim
vc_p8_22_dk
vc_p8_22_nr
vc_p8_22_s0
vc_p8_22_s1
vc_p8_22_s2
vc_p8_23_verbatim
vc_p8_23_dk
vc_p8_23_nr
vc_p8_23_s0
vc_p8_23_s1
vc_p8_23_s2
vc_p8_24_verbatim
vc_p8_24_dk
vc_p8_24_nr
vc_p8_24_s0
vc_p8_24_s1
vc_p8_24_s2
vc_et_raw
fw_st_raw
fw_l_SA_time
fw_l_SA_r1
fw_l_SA_r2
fw_l_SA_r3
fw_l_SA_r4
fw_l_SA_r5
fw_l_SA_dk
fw_l_SA_nr
fw_l_SA_s0
fw_l_SA_s1
fw_l_1_time
fw_l_1_r1
fw_l_1_r2
fw_l_1_r3
fw_l_1_r4
fw_l_1_r5
fw_l_1_dk
fw_l_1_nr
fw_l_1_s0
fw_l_1_s1
fw_l_2_time
fw_l_2_r1
fw_l_2_r2
fw_l_2_r3
fw_l_2_r4
fw_l_2_r5
fw_l_2_dk
fw_l_2_nr
fw_l_2_s0
fw_l_2_s1
fw_l_SB_time
fw_l_SB_r1
fw_l_SB_r2
fw_l_SB_r3
fw_l_SB_r4
fw_l_SB_r5
fw_l_SB_dk
fw_l_SB_nr
fw_l_SB_s0
fw_l_SB_s1
fw_l_3_time
fw_l_3_r1
fw_l_3_r2
fw_l_3_r3
fw_l_3_r4
fw_l_3_r5
fw_l_3_dk
fw_l_3_nr
fw_l_3_s0
fw_l_3_s1
fw_l_4_time
fw_l_4_r1
fw_l_4_r2
fw_l_4_r3
fw_l_4_r4
fw_l_4_r5
fw_l_4_dk
fw_l_4_nr
fw_l_4_s0
fw_l_4_s1
fw_l_5_time
fw_l_5_r1
fw_l_5_r2
fw_l_5_r3
fw_l_5_r4
fw_l_5_r5
fw_l_5_dk
fw_l_5_nr
fw_l_5_s0
fw_l_5_s1
fw_l_6_time
fw_l_6_r1
fw_l_6_r2
fw_l_6_r3
fw_l_6_r4
fw_l_6_r5
fw_l_6_dk
fw_l_6_nr
fw_l_6_s0
fw_l_6_s1
fw_l_7_time
fw_l_7_r1
fw_l_7_r2
fw_l_7_r3
fw_l_7_r4
fw_l_7_r5
fw_l_7_dk
fw_l_7_nr
fw_l_7_s0
fw_l_7_s1
fw_l_8_time
fw_l_8_r1
fw_l_8_r2
fw_l_8_r3
fw_l_8_r4
fw_l_8_r5
fw_l_8_dk
fw_l_8_nr
fw_l_8_s0
fw_l_8_s1
fw_l_9_time
fw_l_9_r1
fw_l_9_r2
fw_l_9_r3
fw_l_9_r4
fw_l_9_r5
fw_l_9_dk
fw_l_9_nr
fw_l_9_s0
fw_l_9_s1
fw_l_10_time
fw_l_10_r1
fw_l_10_r2
fw_l_10_r3
fw_l_10_r4
fw_l_10_r5
fw_l_10_dk
fw_l_10_nr
fw_l_10_s0
fw_l_10_s1
fw_l_11_time
fw_l_11_r1
fw_l_11_r2
fw_l_11_r3
fw_l_11_r4
fw_l_11_r5
fw_l_11_dk
fw_l_11_nr
fw_l_11_s0
fw_l_11_s1
fw_l_12_time
fw_l_12_r1
fw_l_12_r2
fw_l_12_r3
fw_l_12_r4
fw_l_12_r5
fw_l_12_dk
fw_l_12_nr
fw_l_12_s0
fw_l_12_s1
fw_l_13_time
fw_l_13_r1
fw_l_13_r2
fw_l_13_r3
fw_l_13_r4
fw_l_13_r5
fw_l_13_dk
fw_l_13_nr
fw_l_13_s0
fw_l_13_s1
fw_r_14_time
fw_r_14_r1
fw_r_14_r2
fw_r_14_r3
fw_r_14_r4
fw_r_14_r5
fw_r_14_dk
fw_r_14_nr
fw_r_14_s0
fw_r_14_s1_2
fw_r_15_time
fw_r_15_r1
fw_r_15_r2
fw_r_15_r3
fw_r_15_r4
fw_r_15_r5
fw_r_15_dk
fw_r_15_nr
fw_r_15_s0
fw_r_15_s1_2
fw_r_16_time
fw_r_16_r1
fw_r_16_r2
fw_r_16_r3
fw_r_16_r4
fw_r_16_r5
fw_r_16_dk
fw_r_16_nr
fw_r_16_s0
fw_r_16_s1_2
fw_r_17_time
fw_r_17_r1
fw_r_17_r2
fw_r_17_r3
fw_r_17_r4
fw_r_17_r5
fw_r_17_dk
fw_r_17_nr
fw_r_17_s0
fw_r_17_s1_2
fw_r_18_time
fw_r_18_r1
fw_r_18_r2
fw_r_18_r3
fw_r_18_r4
fw_r_18_r5
fw_r_18_dk
fw_r_18_nr
fw_r_18_s0
fw_r_18_s1_2
fw_r_19_time
fw_r_19_r1
fw_r_19_r2
fw_r_19_r3
fw_r_19_r4
fw_r_19_r5
fw_r_19_dk
fw_r_19_nr
fw_r_19_s0
fw_r_19_s1_2
fw_r_20_time
fw_r_20_r1
fw_r_20_r2
fw_r_20_r3
fw_r_20_r4
fw_r_20_r5
fw_r_20_dk
fw_r_20_nr
fw_r_20_s0
fw_r_20_s1_2
fw_r_21_time
fw_r_21_r1
fw_r_21_r2
fw_r_21_r3
fw_r_21_r4
fw_r_21_r5
fw_r_21_dk
fw_r_21_nr
fw_r_21_s0
fw_r_21_s1_2
fw_r_22_time
fw_r_22_r1
fw_r_22_r2
fw_r_22_r3
fw_r_22_r4
fw_r_22_r5
fw_r_22_dk
fw_r_22_nr
fw_r_22_s0
fw_r_22_s1_2
fw_r_23_time
fw_r_23_r1
fw_r_23_r2
fw_r_23_r3
fw_r_23_r4
fw_r_23_r5
fw_r_23_dk
fw_r_23_nr
fw_r_23_s0
fw_r_23_s1_2
fw_r_24_time
fw_r_24_r1
fw_r_24_r2
fw_r_24_r3
fw_r_24_r4
fw_r_24_r5
fw_r_24_dk
fw_r_24_nr
fw_r_24_s0
fw_r_24_s1_2
fw_r_25_time
fw_r_25_r1
fw_r_25_r2
fw_r_25_r3
fw_r_25_r4
fw_r_25_r5
fw_r_25_dk
fw_r_25_nr
fw_r_25_s0
fw_r_25_s1_2
fw_r_26_time
fw_r_26_r1
fw_r_26_r2
fw_r_26_r3
fw_r_26_r4
fw_r_26_r5
fw_r_26_dk
fw_r_26_nr
fw_r_26_s0
fw_r_26_s1_2
fw_r_27_time
fw_r_27_r1
fw_r_27_r2
fw_r_27_r3
fw_r_27_r4
fw_r_27_r5
fw_r_27_dk
fw_r_27_nr
fw_r_27_s0
fw_r_27_s1_2
fw_r_28_time
fw_r_28_r1
fw_r_28_r2
fw_r_28_r3
fw_r_28_r4
fw_r_28_r5
fw_r_28_dk
fw_r_28_nr
fw_r_28_s0
fw_r_28_s1_2
fw_et_raw
vp_st_raw
vp_l_d_time
vp_l_d_r1
vp_l_d_r2
vp_l_d_r3
vp_l_d_r4
vp_l_d_r5
vp_l_d_r6
vp_l_d_dk
vp_l_d_nr
vp_l_d_s0
vp_l_d_s1
vp_l_s_time
vp_l_s_r1
vp_l_s_r2
vp_l_s_r3
vp_l_s_r4
vp_l_s_r5
vp_l_s_r6
vp_l_s_dk
vp_l_s_nr
vp_l_s_s0
vp_l_s_s1
vp_l_1_time
vp_l_1_r1
vp_l_1_r2
vp_l_1_r3
vp_l_1_r4
vp_l_1_r5
vp_l_1_r6
vp_l_1_dk
vp_l_1_nr
vp_l_1_s0
vp_l_1_s1
vp_l_2_time
vp_l_2_r1
vp_l_2_r2
vp_l_2_r3
vp_l_2_r4
vp_l_2_r5
vp_l_2_r6
vp_l_2_dk
vp_l_2_nr
vp_l_2_s0
vp_l_2_s1
vp_l_3_time
vp_l_3_r1
vp_l_3_r2
vp_l_3_r3
vp_l_3_r4
vp_l_3_r5
vp_l_3_r6
vp_l_3_dk
vp_l_3_nr
vp_l_3_s0
vp_l_3_s1
vp_l_4_time
vp_l_4_r1
vp_l_4_r2
vp_l_4_r3
vp_l_4_r4
vp_l_4_r5
vp_l_4_r6
vp_l_4_dk
vp_l_4_nr
vp_l_4_s0
vp_l_4_s1
vp_l_5_time
vp_l_5_r1
vp_l_5_r2
vp_l_5_r3
vp_l_5_r4
vp_l_5_r5
vp_l_5_r6
vp_l_5_dk
vp_l_5_nr
vp_l_5_s0
vp_l_5_s1
vp_l_6_time
vp_l_6_r1
vp_l_6_r2
vp_l_6_r3
vp_l_6_r4
vp_l_6_r5
vp_l_6_r6
vp_l_6_dk
vp_l_6_nr
vp_l_6_s0
vp_l_6_s1
vp_l_7_time
vp_l_7_r1
vp_l_7_r2
vp_l_7_r3
vp_l_7_r4
vp_l_7_r5
vp_l_7_r6
vp_l_7_dk
vp_l_7_nr
vp_l_7_s0
vp_l_7_s1
vp_l_8_time
vp_l_8_r1
vp_l_8_r2
vp_l_8_r3
vp_l_8_r4
vp_l_8_r5
vp_l_8_r6
vp_l_8_dk
vp_l_8_nr
vp_l_8_s0
vp_l_8_s1
vp_l_9_time
vp_l_9_r1
vp_l_9_r2
vp_l_9_r3
vp_l_9_r4
vp_l_9_r5
vp_l_9_r6
vp_l_9_dk
vp_l_9_nr
vp_l_9_s0
vp_l_9_s1
vp_l_10_time
vp_l_10_r1
vp_l_10_r2
vp_l_10_r3
vp_l_10_r4
vp_l_10_r5
vp_l_10_r6
vp_l_10_dk
vp_l_10_nr
vp_l_10_s0
vp_l_10_s1
vp_l_11_time
vp_l_11_r1
vp_l_11_r2
vp_l_11_r3
vp_l_11_r4
vp_l_11_r5
vp_l_11_r6
vp_l_11_dk
vp_l_11_nr
vp_l_11_s0
vp_l_11_s1
vp_l_12_time
vp_l_12_r1
vp_l_12_r2
vp_l_12_r3
vp_l_12_r4
vp_l_12_r5
vp_l_12_r6
vp_l_12_dk
vp_l_12_nr
vp_l_12_s0
vp_l_12_s1
vp_r_13_time
vp_r_13_r1
vp_r_13_r2
vp_r_13_r3
vp_r_13_r4
vp_r_13_r5
vp_r_13_r6
vp_r_13_dk
vp_r_13_nr
vp_r_13_s0
vp_r_13_s1
vp_r_14_time
vp_r_14_r1
vp_r_14_r2
vp_r_14_r3
vp_r_14_r4
vp_r_14_r5
vp_r_14_r6
vp_r_14_dk
vp_r_14_nr
vp_r_14_s0
vp_r_14_s1
vp_r_15_time
vp_r_15_r1
vp_r_15_r2
vp_r_15_r3
vp_r_15_r4
vp_r_15_r5
vp_r_15_r6
vp_r_15_dk
vp_r_15_nr
vp_r_15_s0
vp_r_15_s1
vp_r_16_time
vp_r_16_r1
vp_r_16_r2
vp_r_16_r3
vp_r_16_r4
vp_r_16_r5
vp_r_16_r6
vp_r_16_dk
vp_r_16_nr
vp_r_16_s0
vp_r_16_s1
vp_r_17_time
vp_r_17_r1
vp_r_17_r2
vp_r_17_r3
vp_r_17_r4
vp_r_17_r5
vp_r_17_r6
vp_r_17_dk
vp_r_17_nr
vp_r_17_s0
vp_r_17_s1
vp_r_18_time
vp_r_18_r1
vp_r_18_r2
vp_r_18_r3
vp_r_18_r4
vp_r_18_r5
vp_r_18_r6
vp_r_18_dk
vp_r_18_nr
vp_r_18_s0
vp_r_18_s1
vp_r_19_time
vp_r_19_r1
vp_r_19_r2
vp_r_19_r3
vp_r_19_r4
vp_r_19_r5
vp_r_19_r6
vp_r_19_dk
vp_r_19_nr
vp_r_19_s0
vp_r_19_s1
vp_r_20_time
vp_r_20_r1
vp_r_20_r2
vp_r_20_r3
vp_r_20_r4
vp_r_20_r5
vp_r_20_r6
vp_r_20_dk
vp_r_20_nr
vp_r_20_s0
vp_r_20_s1
vp_r_21_time
vp_r_21_r1
vp_r_21_r2
vp_r_21_r3
vp_r_21_r4
vp_r_21_r5
vp_r_21_r6
vp_r_21_dk
vp_r_21_nr
vp_r_21_s0
vp_r_21_s1
vp_r_22_time
vp_r_22_r1
vp_r_22_r2
vp_r_22_r3
vp_r_22_r4
vp_r_22_r5
vp_r_22_r6
vp_r_22_dk
vp_r_22_nr
vp_r_22_s0
vp_r_22_s1
vp_r_23_time
vp_r_23_r1
vp_r_23_r2
vp_r_23_r3
vp_r_23_r4
vp_r_23_r5
vp_r_23_r6
vp_r_23_dk
vp_r_23_nr
vp_r_23_s0
vp_r_23_s1
vp_r_24_time
vp_r_24_r1
vp_r_24_r2
vp_r_24_r3
vp_r_24_r4
vp_r_24_r5
vp_r_24_r6
vp_r_24_dk
vp_r_24_nr
vp_r_24_s0
vp_r_24_s1
vp_r_25_time
vp_r_25_r1
vp_r_25_r2
vp_r_25_r3
vp_r_25_r4
vp_r_25_r5
vp_r_25_r6
vp_r_25_dk
vp_r_25_nr
vp_r_25_s0
vp_r_25_s1
vp_et_raw
rd_st_raw
rd_1_verbatim
rd_1_s
rd_2_verbatim
rd_2_s
rd_3_verbatim
rd_3_s
rd_4_verbatim
rd_4_s
rd_5_verbatim
rd_5_s
rd_ra_verbatim
rd_ra_s
rd_6_verbatim
rd_6_s
rd_7_verbatim
rd_7_s
rd_8_verbatim
rd_8_s
rd_rb_verbatim
rd_rb_s
rd_9_verbatim
rd_9_s
rd_10_verbatim
rd_10_s
rd_et_raw
ss_st_raw
ss_t
ss_et_raw
in_st_raw
in_p12_1_verbatim
in_p12_1_dk
in_p12_1_nr
in_p12_1_s0
in_p12_1_s1
in_p12_2_verbatim
in_p12_2_dk
in_p12_2_nr
in_p12_2_s0
in_p12_2_s1
in_p12_3_verbatim
in_p12_3_dk
in_p12_3_nr
in_p12_3_s0
in_p12_3_s1
in_p12_4_verbatim
in_p12_4_dk
in_p12_4_nr
in_p12_4_s0
in_p12_4_s1
in_p12_5_verbatim
in_p12_5_dk
in_p12_5_nr
in_p12_5_s0
in_p12_5_s1
in_p12_6_verbatim
in_p12_6_dk
in_p12_6_nr
in_p12_6_s0
in_p12_6_s1
in_p12_7_verbatim
in_p12_7_dk
in_p12_7_nr
in_p12_7_s0
in_p12_7_s1
in_p12_8_verbatim
in_p12_8_dk
in_p12_8_nr
in_p12_8_s0
in_p12_8_s1
in_p12_9_verbatim
in_p12_9_dk
in_p12_9_nr
in_p12_9_s0
in_p12_9_s1
in_p13_10_Verbatim
in_p13_10_dk
in_p13_10_nr
in_p13_10_s0
in_p13_10_s1
in_p13_11_Verbatim
in_p13_11_dk
in_p13_11_nr
in_p13_11_s0
in_p13_11_s1
in_p13_12_Verbatim
in_p13_12_dk
in_p13_12_nr
in_p13_12_s0
in_p13_12_s1
in_p13_13_Verbatim
in_p13_13_dk
in_p13_13_nr
in_p13_13_s0
in_p13_13_s1
in_p13_14_Verbatim
in_p13_14_dk
in_p13_14_nr
in_p13_14_s0
in_p13_14_s1
in_p13_15_Verbatim
in_p13_15_dk
in_p13_15_nr
in_p13_15_s0
in_p13_15_s1
in_p13_16_Verbatim
in_p13_16_dk
in_p13_16_nr
in_p13_16_s0
in_p13_16_s1
in_p13_17_Verbatim
in_p13_17_dk
in_p13_17_nr
in_p13_17_s0
in_p13_17_s1
in_p13_18_Verbatim
in_p13_18_dk
in_p13_18_nr
in_p13_18_s0
in_p13_18_s1
in_p13_19_Verbatim
in_p13_19_dk
in_p13_19_nr
in_p13_19_s0
in_p13_19_s1
in_p13_20_Verbatim
in_p13_20_dk
in_p13_20_nr
in_p13_20_s0
in_p13_20_s1
in_p13_21_Verbatim
in_p13_21_dk
in_p13_21_nr
in_p13_21_s0
in_p13_21_s1
in_p13_22_Verbatim
in_p13_22_dk
in_p13_22_nr
in_p13_22_s0
in_p13_22_s1
in_p13_23_Verbatim
in_p13_23_dk
in_p13_23_nr
in_p13_23_s0
in_p13_23_s1
in_p13_24_Verbatim
in_p13_24_dk
in_p13_24_nr
in_p13_24_s0
in_p13_24_s1
in_et_raw
ar_st_raw
ar_l_1_time
ar_l_1_verbatim
ar_l_1_dk
ar_l_1_nr
ar_l_1_ir
ar_l_1_s0
ar_l_1_s1
ar_l_2_time
ar_l_2_verbatim
ar_l_2_dk
ar_l_2_nr
ar_l_2_ir
ar_l_2_s0
ar_l_2_s1
ar_l_3_time
ar_l_3_verbatim
ar_l_3_dk
ar_l_3_nr
ar_l_3_ir
ar_l_3_s0
ar_l_3_s1
ar_l_s_time
ar_l_s_verbatim
ar_l_s_dk
ar_l_s_nr
ar_l_s_ir
ar_l_s_s0
ar_l_s_s1
ar_l_4_time
ar_l_4_verbatim
ar_l_4_dk
ar_l_4_nr
ar_l_4_ir
ar_l_4_s0
ar_l_4_s1
ar_l_5_time
ar_l_5_verbatim
ar_l_5_dk
ar_l_5_nr
ar_l_5_ir
ar_l_5_s0
ar_l_5_s1
ar_l_6_time
ar_l_6_verbatim
ar_l_6_dk
ar_l_6_nr
ar_l_6_ir
ar_l_6_s0
ar_l_6_s1
ar_l_7_time
ar_l_7_verbatim
ar_l_7_dk
ar_l_7_nr
ar_l_7_ir
ar_l_7_s0
ar_l_7_s1
ar_l_8_time
ar_l_8_verbatim
ar_l_8_dk
ar_l_8_nr
ar_l_8_ir
ar_l_8_s0
ar_l_8_s1
ar_l_9_time
ar_l_9_verbatim
ar_l_9_dk
ar_l_9_nr
ar_l_9_ir
ar_l_9_s0
ar_l_9_s1
ar_l_10_time
ar_l_10_verbatim
ar_l_10_dk
ar_l_10_nr
ar_l_10_ir
ar_l_10_s0
ar_l_10_s1
ar_l_11_time
ar_l_11_verbatim
ar_l_11_dk
ar_l_11_nr
ar_l_11_ir
ar_l_11_s0
ar_l_11_s1
ar_l_12_time
ar_l_12_verbatim
ar_l_12_dk
ar_l_12_nr
ar_l_12_ir
ar_l_12_s0
ar_l_12_s1
ar_r_13_time
ar_r_13_verbatim
ar_r_13_dk
ar_r_13_nr
ar_r_13_ir
ar_r_13_s0
ar_r_13_s1
ar_r_13_s2
ar_r_14_time
ar_r_14_verbatim
ar_r_14_dk
ar_r_14_nr
ar_r_14_ir
ar_r_14_s0
ar_r_14_s1
ar_r_14_s2
ar_r_15_time
ar_r_15_verbatim
ar_r_15_dk
ar_r_15_nr
ar_r_15_ir
ar_r_15_s0
ar_r_15_s1
ar_r_15_s2
ar_r_16_time
ar_r_16_verbatim
ar_r_16_dk
ar_r_16_nr
ar_r_16_ir
ar_r_16_s0
ar_r_16_s1
ar_r_16_s2
ar_r_17_time
ar_r_17_verbatim
ar_r_17_dk
ar_r_17_nr
ar_r_17_ir
ar_r_17_s0
ar_r_17_s1
ar_r_17_s2
ar_r_18_time
ar_r_18_verbatim
ar_r_18_dk
ar_r_18_nr
ar_r_18_ir
ar_r_18_s0
ar_r_18_s1
ar_r_18_s2
ar_r_19_time
ar_r_19_verbatim
ar_r_19_dk
ar_r_19_nr
ar_r_19_ir
ar_r_19_s0
ar_r_19_s1
ar_r_19_s2
ar_r_20_time
ar_r_20_verbatim
ar_r_20_dk
ar_r_20_nr
ar_r_20_ir
ar_r_20_s0
ar_r_20_s1
ar_r_20_s2
ar_r_21_time
ar_r_21_verbatim
ar_r_21_dk
ar_r_21_nr
ar_r_21_ir
ar_r_21_s0
ar_r_21_s1
ar_r_21_s2
ar_r_22_time
ar_r_22_verbatim
ar_r_22_dk
ar_r_22_nr
ar_r_22_ir
ar_r_22_s0
ar_r_22_s1
ar_r_22_s2
ar_et_raw
db_st_raw
db_p14_s_1_verbatim
db_p14_s_1_dk
db_p14_s_1_nr
db_p14_s_1_t0
db_p14_s_1_t1
db_p14_s_2_verbatim
db_p14_s_2_dk
db_p14_s_2_nr
db_p14_s_2_t0
db_p14_s_2_t1
db_p14_1_1_verbatim
db_p14_1_1_dk
db_p14_1_1_nr
db_p14_1_1_t0
db_p14_1_1_t1
db_p14_1_2_verbatim
db_p14_1_2_dk
db_p14_1_2_nr
db_p14_1_2_t0
db_p14_1_2_t1
db_p14_2_1_verbatim
db_p14_2_1_dk
db_p14_2_1_nr
db_p14_2_1_t0
db_p14_2_1_t1
db_p14_2_2_verbatim
db_p14_2_2_dk
db_p14_2_2_nr
db_p14_2_2_t0
db_p14_2_2_t1
db_p14_3_1_verbatim
db_p14_3_1_dk
db_p14_3_1_nr
db_p14_3_1_t0
db_p14_3_1_t1
db_p14_3_2_verbatim
db_p14_3_2_dk
db_p14_3_2_nr
db_p14_3_2_t0
db_p14_3_2_t1
db_p14_scores_1_s0
db_p14_scores_1_s1
db_p14_scores_1_s2
db_p14_scores_2_s0
db_p14_scores_2_s1
db_p14_scores_2_s2
db_p14_scores_3_s0
db_p14_scores_3_s1
db_p14_scores_3_s2
db_p15_4_1_verbatim
db_p15_4_1_dk
db_p15_4_1_nr
db_p15_4_1_t0
db_p15_4_1_t1_2
db_p15_4_2_verbatim
db_p15_4_2_dk
db_p15_4_2_nr
db_p15_4_2_t0
db_p15_4_2_t1_2
db_p15_5_1_verbatim
db_p15_5_1_dk
db_p15_5_1_nr
db_p15_5_1_t0
db_p15_5_1_t1_2
db_p15_5_2_verbatim
db_p15_5_2_dk
db_p15_5_2_nr
db_p15_5_2_t0
db_p15_5_2_t1_2
db_p15_6_1_verbatim
db_p15_6_1_dk
db_p15_6_1_nr
db_p15_6_1_t0
db_p15_6_1_t1_2
db_p15_6_2_verbatim
db_p15_6_2_dk
db_p15_6_2_nr
db_p15_6_2_t0
db_p15_6_2_t1_2
db_p15_7_1_verbatim
db_p15_7_1_dk
db_p15_7_1_nr
db_p15_7_1_t0
db_p15_7_1_t1_2
db_p15_7_2_verbatim
db_p15_7_2_dk
db_p15_7_2_nr
db_p15_7_2_t0
db_p15_7_2_t1_2
db_p15_8_1_verbatim
db_p15_8_1_dk
db_p15_8_1_nr
db_p15_8_1_t0
db_p15_8_1_t1_2
db_p15_8_2_verbatim
db_p15_8_2_dk
db_p15_8_2_nr
db_p15_8_2_t0
db_p15_8_2_t1_2
db_p15_9_1_verbatim
db_p15_9_1_dk
db_p15_9_1_nr
db_p15_9_1_t0
db_p15_9_1_t1_2
db_p15_9_2_verbatim
db_p15_9_2_dk
db_p15_9_2_nr
db_p15_9_2_t0
db_p15_9_2_t1_2
db_p15_10_1_verbatim
db_p15_10_1_dk
db_p15_10_1_nr
db_p15_10_1_t0
db_p15_10_1_t1_2
db_p15_10_2_verbatim
db_p15_10_2_dk
db_p15_10_2_nr
db_p15_10_2_t0
db_p15_10_2_t1_2
db_p15_scores_4_s0
db_p15_scores_4_s1_2
db_p15_scores_4_s2_4
db_p15_scores_5_s0
db_p15_scores_5_s1_2
db_p15_scores_5_s2_4
db_p15_scores_6_s0
db_p15_scores_6_s1_2
db_p15_scores_6_s2_4
db_p15_scores_7_s0
db_p15_scores_7_s1_2
db_p15_scores_7_s2_4
db_p15_scores_8_s0
db_p15_scores_8_s1_2
db_p15_scores_8_s2_4
db_p15_scores_9_s0
db_p15_scores_9_s1_2
db_p15_scores_9_s2_4
db_p15_scores_10_s0
db_p15_scores_10_s1_2
db_p15_scores_10_s2_4
db_et_raw
ssp_st_raw
ssp_p15_sa_verbatim
ssp_p15_sa_dk
ssp_p15_sa_nr
ssp_p15_sa_s0
ssp_p15_sa_s1
ssp_p15_sa_s2
ssp_p15_1_verbatim
ssp_p15_1_dk
ssp_p15_1_nr
ssp_p15_1_s0
ssp_p15_1_s1
ssp_p15_1_s2
ssp_p15_2_verbatim
ssp_p15_2_dk
ssp_p15_2_nr
ssp_p15_2_s0
ssp_p15_2_s1
ssp_p15_2_s2
ssp_p15_3_verbatim
ssp_p15_3_dk
ssp_p15_3_nr
ssp_p15_3_s0
ssp_p15_3_s1
ssp_p15_3_s2
ssp_p15_4_verbatim
ssp_p15_4_dk
ssp_p15_4_nr
ssp_p15_4_s0
ssp_p15_4_s1
ssp_p15_4_s2
ssp_p15_5_verbatim
ssp_p15_5_dk
ssp_p15_5_nr
ssp_p15_5_s0
ssp_p15_5_s1
ssp_p15_5_s2
ssp_p15_6_verbatim
ssp_p15_6_dk
ssp_p15_6_nr
ssp_p15_6_s0
ssp_p15_6_s1
ssp_p15_6_s2
ssp_p15_7_verbatim
ssp_p15_7_dk
ssp_p15_7_nr
ssp_p15_7_s0
ssp_p15_7_s1
ssp_p15_7_s2
ssp_p15_8_verbatim
ssp_p15_8_dk
ssp_p15_8_nr
ssp_p15_8_s0
ssp_p15_8_s1
ssp_p15_8_s2
ssp_p15_9_verbatim
ssp_p15_9_dk
ssp_p15_9_nr
ssp_p15_9_s0
ssp_p15_9_s1
ssp_p15_9_s2
ssp_p16_10_verbatim
ssp_p16_10_dk
ssp_p16_10_nr
ssp_p16_10_s0
ssp_p16_10_s1
ssp_p16_10_s2
ssp_p16_11_verbatim
ssp_p16_11_dk
ssp_p16_11_nr
ssp_p16_11_s0
ssp_p16_11_s1
ssp_p16_11_s2
ssp_p16_12_verbatim
ssp_p16_12_dk
ssp_p16_12_nr
ssp_p16_12_s0
ssp_p16_12_s1
ssp_p16_12_s2
ssp_p16_13_verbatim
ssp_p16_13_dk
ssp_p16_13_nr
ssp_p16_13_s0
ssp_p16_13_s1
ssp_p16_13_s2
ssp_p16_14_verbatim
ssp_p16_14_dk
ssp_p16_14_nr
ssp_p16_14_s0
ssp_p16_14_s1
ssp_p16_14_s2
ssp_p16_15_verbatim
ssp_p16_15_dk
ssp_p16_15_nr
ssp_p16_15_s0
ssp_p16_15_s1
ssp_p16_15_s2
ssp_p16_16_verbatim
ssp_p16_16_dk
ssp_p16_16_nr
ssp_p16_16_s0
ssp_p16_16_s1
ssp_p16_16_s2
ssp_p16_17_verbatim
ssp_p16_17_dk
ssp_p16_17_nr
ssp_p16_17_s0
ssp_p16_17_s1
ssp_p16_17_s2
ssp_p16_18_verbatim
ssp_p16_18_dk
ssp_p16_18_nr
ssp_p16_18_s0
ssp_p16_18_s1
ssp_p16_18_s2
ssp_p16_19_verbatim
ssp_p16_19_dk
ssp_p16_19_nr
ssp_p16_19_s0
ssp_p16_19_s1
ssp_p16_19_s2
ssp_p16_20_verbatim
ssp_p16_20_dk
ssp_p16_20_nr
ssp_p16_20_s0
ssp_p16_20_s1
ssp_p16_20_s2
ssp_p16_21_verbatim
ssp_p16_21_dk
ssp_p16_21_nr
ssp_p16_21_s0
ssp_p16_21_s1
ssp_p16_21_s2
ssp_p16_22_verbatim
ssp_p16_22_dk
ssp_p16_22_nr
ssp_p16_22_s0
ssp_p16_22_s1
ssp_p16_22_s2
ssp_p16_23_verbatim
ssp_p16_23_dk
ssp_p16_23_nr
ssp_p16_23_s0
ssp_p16_23_s1
ssp_p16_23_s2
ssp_et_raw
nsq_st_raw
nsq_t1
nsq_t2


nsq_et_raw
co_st_raw
co_p17_1_verbatim
co_p17_1_dk
co_p17_1_nr
co_p17_1_s0
co_p17_1_s1
co_p17_1_s2
co_p17_2_verbatim
co_p17_2_dk
co_p17_2_nr
co_p17_2_s0
co_p17_2_s1
co_p17_2_s2
co_p17_3_verbatim
co_p17_3_dk
co_p17_3_nr
co_p17_3_s0
co_p17_3_s1
co_p17_3_s2
co_p17_4_verbatim
co_p17_4_dk
co_p17_4_nr
co_p17_4_s0
co_p17_4_s1
co_p17_4_s2
co_p17_5_verbatim
co_p17_5_dk
co_p17_5_nr
co_p17_5_s0
co_p17_5_s1
co_p17_5_s2
co_p17_6_verbatim
co_p17_6_dk
co_p17_6_nr
co_p17_6_s0
co_p17_6_s1
co_p17_6_s2
co_p17_7_verbatim
co_p17_7_dk
co_p17_7_nr
co_p17_7_s0
co_p17_7_s1
co_p17_7_s2
co_p17_8_verbatim
co_p17_8_dk
co_p17_8_nr
co_p17_8_s0
co_p17_8_s1
co_p17_8_s2
co_p17_9_verbatim
co_p17_9_dk
co_p17_9_nr
co_p17_9_s0
co_p17_9_s1
co_p17_9_s2
co_p18_10_verbatim
co_p18_10_dk
co_p18_10_nr
co_p18_10_s0
co_p18_10_s1
co_p18_10_s2
co_p18_11_verbatim
co_p18_11_dk
co_p18_11_nr
co_p18_11_s0
co_p18_11_s1
co_p18_11_s2
co_p18_12_verbatim
co_p18_12_dk
co_p18_12_nr
co_p18_12_s0
co_p18_12_s1
co_p18_12_s2
co_p18_13_verbatim
co_p18_13_dk
co_p18_13_nr
co_p18_13_s0
co_p18_13_s1
co_p18_13_s2
co_p18_14_verbatim
co_p18_14_dk
co_p18_14_nr
co_p18_14_s0
co_p18_14_s1
co_p18_14_s2
co_p18_15_verbatim
co_p18_15_dk
co_p18_15_nr
co_p18_15_s0
co_p18_15_s1
co_p18_15_s2
co_p18_16_verbatim
co_p18_16_dk
co_p18_16_nr
co_p18_16_s0
co_p18_16_s1
co_p18_16_s2
co_p18_17_verbatim
co_p18_17_dk
co_p18_17_nr
co_p18_17_s0
co_p18_17_s1
co_p18_17_s2
co_p18_18_verbatim
co_p18_18_dk
co_p18_18_nr
co_p18_18_s0
co_p18_18_s1
co_p18_18_s2
co_et_raw
sr_st_raw
sr_p19_1_r1
sr_p19_1_r2
sr_p19_1_r3
sr_p19_1_r4
sr_p19_1_r5
sr_p19_1_dk
sr_p19_1_nr
sr_p19_1_s0
sr_p19_1_s1
sr_p19_2_r1
sr_p19_2_r2
sr_p19_2_r3
sr_p19_2_r4
sr_p19_2_r5
sr_p19_2_dk
sr_p19_2_nr
sr_p19_2_s0
sr_p19_2_s1
sr_p19_3_r1
sr_p19_3_r2
sr_p19_3_r3
sr_p19_3_r4
sr_p19_3_r5
sr_p19_3_dk
sr_p19_3_nr
sr_p19_3_s0
sr_p19_3_s1
sr_p19_4_r1
sr_p19_4_r2
sr_p19_4_r3
sr_p19_4_r4
sr_p19_4_r5
sr_p19_4_dk
sr_p19_4_nr
sr_p19_4_s0
sr_p19_4_s1
sr_p19_DA_r1
sr_p19_DA_r2
sr_p19_DA_r3
sr_p19_DA_r4
sr_p19_DA_r5
sr_p19_DA_dk
sr_p19_DA_nr
sr_p19_DA_s0
sr_p19_DA_s1
sr_p19_DB_r1
sr_p19_DB_r2
sr_p19_DB_r3
sr_p19_DB_r4
sr_p19_DB_r5
sr_p19_DB_dk
sr_p19_DB_nr
sr_p19_DB_s0
sr_p19_DB_s1
sr_p19_DC_r1
sr_p19_DC_r2
sr_p19_DC_r3
sr_p19_DC_r4
sr_p19_DC_r5
sr_p19_DC_dk
sr_p19_DC_nr
sr_p19_DC_s0
sr_p19_DC_s1
sr_p19_5_r1
sr_p19_5_r2
sr_p19_5_r3
sr_p19_5_r4
sr_p19_5_r5
sr_p19_5_dk
sr_p19_5_nr
sr_p19_5_s0
sr_p19_5_s1
sr_p19_6_r1
sr_p19_6_r2
sr_p19_6_r3
sr_p19_6_r4
sr_p19_6_r5
sr_p19_6_dk
sr_p19_6_nr
sr_p19_6_s0
sr_p19_6_s1
sr_p19_7_r1
sr_p19_7_r2
sr_p19_7_r3
sr_p19_7_r4
sr_p19_7_r5
sr_p19_7_dk
sr_p19_7_nr
sr_p19_7_s0
sr_p19_7_s1
sr_p19_8_r1
sr_p19_8_r2
sr_p19_8_r3
sr_p19_8_r4
sr_p19_8_r5
sr_p19_8_dk
sr_p19_8_nr
sr_p19_8_s0
sr_p19_8_s1
sr_p19_9_r1
sr_p19_9_r2
sr_p19_9_r3
sr_p19_9_r4
sr_p19_9_r5
sr_p19_9_dk
sr_p19_9_nr
sr_p19_9_s0
sr_p19_9_s1
sr_p19_10_r1
sr_p19_10_r2
sr_p19_10_r3
sr_p19_10_r4
sr_p19_10_r5
sr_p19_10_dk
sr_p19_10_nr
sr_p19_10_s0
sr_p19_10_s1
sr_p19_11_r1
sr_p19_11_r2
sr_p19_11_r3
sr_p19_11_r4
sr_p19_11_r5
sr_p19_11_dk
sr_p19_11_nr
sr_p19_11_s0
sr_p19_11_s1
sr_p19_12_r1
sr_p19_12_r2
sr_p19_12_r3
sr_p19_12_r4
sr_p19_12_r5
sr_p19_12_dk
sr_p19_12_nr
sr_p19_12_s0
sr_p19_12_s1
sr_p20_13_r1
sr_p20_13_r2
sr_p20_13_r3
sr_p20_13_r4
sr_p20_13_dk
sr_p20_13_nr
sr_p20_13_s0
sr_p20_13_s1
sr_p20_13_s2
sr_p20_14_r1
sr_p20_14_r2
sr_p20_14_r3
sr_p20_14_r4
sr_p20_14_dk
sr_p20_14_nr
sr_p20_14_s0
sr_p20_14_s1
sr_p20_14_s2
sr_p20_15_r1
sr_p20_15_r2
sr_p20_15_r3
sr_p20_15_r4
sr_p20_15_dk
sr_p20_15_nr
sr_p20_15_s0
sr_p20_15_s1
sr_p20_15_s2
sr_p20_16_r1
sr_p20_16_r2
sr_p20_16_r3
sr_p20_16_r4
sr_p20_16_dk
sr_p20_16_nr
sr_p20_16_s0
sr_p20_16_s1
sr_p20_16_s2
sr_p20_17_r1
sr_p20_17_r2
sr_p20_17_r3
sr_p20_17_r4
sr_p20_17_dk
sr_p20_17_nr
sr_p20_17_s0
sr_p20_17_s1
sr_p20_17_s2
sr_p20_18_r1
sr_p20_18_r2
sr_p20_18_r3
sr_p20_18_r4
sr_p20_18_dk
sr_p20_18_nr
sr_p20_18_s0
sr_p20_18_s1
sr_p20_18_s2
sr_p20_19_r1
sr_p20_19_r2
sr_p20_19_r3
sr_p20_19_r4
sr_p20_19_dk
sr_p20_19_nr
sr_p20_19_s0
sr_p20_19_s1
sr_p20_19_s2
sr_p20_20_r1
sr_p20_20_r2
sr_p20_20_r3
sr_p20_20_r4
sr_p20_20_dk
sr_p20_20_nr
sr_p20_20_s0
sr_p20_20_s1
sr_p20_20_s2
sr_p20_21_r1
sr_p20_21_r2
sr_p20_21_r3
sr_p20_21_r4
sr_p20_21_dk
sr_p20_21_nr
sr_p20_21_s0
sr_p20_21_s1
sr_p20_21_s2
sr_p20_22_r1
sr_p20_22_r2
sr_p20_22_r3
sr_p20_22_r4
sr_p20_22_dk
sr_p20_22_nr
sr_p20_22_s0
sr_p20_22_s1
sr_p20_22_s2
sr_p20_23_r1
sr_p20_23_r2
sr_p20_23_r3
sr_p20_23_r4
sr_p20_23_dk
sr_p20_23_nr
sr_p20_23_s0
sr_p20_23_s1
sr_p20_23_s2
sr_p20_24_r1
sr_p20_24_r2
sr_p20_24_r3
sr_p20_24_r4
sr_p20_24_dk
sr_p20_24_nr
sr_p20_24_s0
sr_p20_24_s1
sr_p20_24_s2
sr_p20_25_r1
sr_p20_25_r2
sr_p20_25_r3
sr_p20_25_r4
sr_p20_25_dk
sr_p20_25_nr
sr_p20_25_s0
sr_p20_25_s1
sr_p20_25_s2
sr_p20_26_r1
sr_p20_26_r2
sr_p20_26_r3
sr_p20_26_r4
sr_p20_26_dk
sr_p20_26_nr
sr_p20_26_s0
sr_p20_26_s1
sr_p20_26_s2
sr_p20_27_r1
sr_p20_27_r2
sr_p20_27_r3
sr_p20_27_r4
sr_p20_27_dk
sr_p20_27_nr
sr_p20_27_s0
sr_p20_27_s1
sr_p20_27_s2
sr_et_raw
sa_st_raw
sa_1_5_1_s0
sa_1_5_1_s1
sa_1_5_2_s0
sa_1_5_2_s1
sa_1_5_3_s0
sa_1_5_3_s1
sa_1_5_4_s0
sa_1_5_4_s1
sa_1_5_5_s0
sa_1_5_5_s1
sa_6_14_6_s0
sa_6_14_6_s1
sa_6_14_7_s0
sa_6_14_7_s1
sa_6_14_8_s0
sa_6_14_8_s1
sa_6_14_9_s0
sa_6_14_9_s1
sa_6_14_10_s0
sa_6_14_10_s1
sa_6_14_11_s0
sa_6_14_11_s1
sa_6_14_12_s0
sa_6_14_12_s1
sa_6_14_13_s0
sa_6_14_13_s1
sa_6_14_14_s0
sa_6_14_14_s1
sa_15_23_15_s0
sa_15_23_15_s1
sa_15_23_16_s0
sa_15_23_16_s1
sa_15_23_17_s0
sa_15_23_17_s1
sa_15_23_18_s0
sa_15_23_18_s1
sa_15_23_19_s0
sa_15_23_19_s1
sa_15_23_20_s0
sa_15_23_20_s1
sa_15_23_21_s0
sa_15_23_21_s1
sa_15_23_22_s0
sa_15_23_22_s1
sa_15_23_23_s0
sa_15_23_23_s1
sa_et_raw
lns_st_raw
lns_d_verbatim
lns_d_dk
lns_d_nr
lns_d_t0
lns_d_t1_2
lns_sa_verbatim
lns_sa_dk
lns_sa_nr
lns_sa_t0
lns_sa_t1_2
lns_1_1_verbatim
lns_1_1_dk
lns_1_1_nr
lns_1_1_s0
lns_1_1_s1_2
lns_1_2_verbatim
lns_1_2_dk
lns_1_2_nr
lns_1_2_s0
lns_1_2_s1_2
lns_2_1_verbatim
lns_2_1_dk
lns_2_1_nr
lns_2_1_s0
lns_2_1_s1_2
lns_2_2_verbatim
lns_2_2_dk
lns_2_2_nr
lns_2_2_s0
lns_2_2_s1_2
lns_sb_1_verbatim
lns_sb_1_dk
lns_sb_1_nr
lns_sb_1_s0
lns_sb_1_s1_2
lns_sb_2_verbatim
lns_sb_2_dk
lns_sb_2_nr
lns_sb_2_s0
lns_sb_2_s1_2
lns_3_1_verbatim
lns_3_1_dk
lns_3_1_nr
lns_3_1_s0
lns_3_1_s1_2
lns_3_2_verbatim
lns_3_2_dk
lns_3_2_nr
lns_3_2_s0
lns_3_2_s1_2
lns_4_1_verbatim
lns_4_1_dk
lns_4_1_nr
lns_4_1_s0
lns_4_1_s1_2
lns_4_2_verbatim
lns_4_2_dk
lns_4_2_nr
lns_4_2_s0
lns_4_2_s1_2
lns_5_1_verbatim
lns_5_1_dk
lns_5_1_nr
lns_5_1_s0
lns_5_1_s1_2
lns_5_2_verbatim
lns_5_2_dk
lns_5_2_nr
lns_5_2_s0
lns_5_2_s1_2
lns_6_1_verbatim
lns_6_1_dk
lns_6_1_nr
lns_6_1_s0
lns_6_1_s1_2
lns_6_2_verbatim
lns_6_2_dk
lns_6_2_nr
lns_6_2_s0
lns_6_2_s1_2
lns_7_1_verbatim
lns_7_1_dk
lns_7_1_nr
lns_7_1_s0
lns_7_1_s1_2
lns_7_2_verbatim
lns_7_2_dk
lns_7_2_nr
lns_7_2_s0
lns_7_2_s1_2
lns_8_1_verbatim
lns_8_1_dk
lns_8_1_nr
lns_8_1_s0
lns_8_1_s1_2
lns_8_2_verbatim
lns_8_2_dk
lns_8_2_nr
lns_8_2_s0
lns_8_2_s1_2
lns_9_1_verbatim
lns_9_1_dk
lns_9_1_nr
lns_9_1_s0
lns_9_1_s1_2
lns_9_2_verbatim
lns_9_2_dk
lns_9_2_nr
lns_9_2_s0
lns_9_2_s1_2
lns_10_1_verbatim
lns_10_1_dk
lns_10_1_nr
lns_10_1_s0
lns_10_1_s1_2
lns_10_2_verbatim
lns_10_2_dk
lns_10_2_nr
lns_10_2_s0
lns_10_2_s1_2
lns_scores_1_s0
lns_scores_1_s1_2
lns_scores_1_s2_4
lns_scores_2_s0
lns_scores_2_s1_2
lns_scores_2_s2_4
lns_scores_3_s0
lns_scores_3_s1_2
lns_scores_3_s2_4
lns_scores_4_s0
lns_scores_4_s1_2
lns_scores_4_s2_4
lns_scores_5_s0
lns_scores_5_s1_2
lns_scores_5_s2_4
lns_scores_6_s0
lns_scores_6_s1_2
lns_scores_6_s2_4
lns_scores_7_s0
lns_scores_7_s1_2
lns_scores_7_s2_4
lns_scores_8_s0
lns_scores_8_s1_2
lns_scores_8_s2_4
lns_scores_9_s0
lns_scores_9_s1_2
lns_scores_9_s2_4
lns_scores_10_s0
lns_scores_10_s1_2
lns_scores_10_s2_4
lns_et_raw
mdf_1
mdf_2
mdf_3
mdf_4
mdf_5
mdf_6
mdf_7
mdf_text
gn
""".strip()


def load_final_column_order():
    cols = [ln.strip() for ln in FINAL_COLUMN_ORDER_TEXT.splitlines() if ln.strip()]
    # De-duplicate while preserving first occurrence
    return list(dict.fromkeys(cols))


def apply_final_column_order(df: pd.DataFrame) -> pd.DataFrame:
    desired = load_final_column_order()

    missing = [c for c in desired if c not in df.columns]
    if missing:
        print(f"⚠️ {len(missing)} columns in FINAL_COLUMN_ORDER_TEXT were not found in the dataframe.")
        print("   First 30 missing:", missing[:30])

    # Enforce desired order; append any remaining columns at the end
    ordered_present = [c for c in desired if c in df.columns]
    remainder = [c for c in df.columns if c not in desired]
    return df.reindex(columns=ordered_present + remainder)


# ============================================================
# === FIELD VALUE EXTRACTION                                ===
# ============================================================
def best_value(field_obj):
    """
    Extract a usable value from a DI field object.
    Prefers: content -> valueString -> other value* keys.
    NOTE: This does not normalise "selected"/"None" etc.
    """
    if not isinstance(field_obj, dict):
        return field_obj

    val = field_obj.get("content") or field_obj.get("valueString")

    if val is None:
        for k in (
            "value", "valueNumber", "valueInteger", "valueDate", "valueTime",
            "valuePhoneNumber", "valueBoolean", "valueSelectionGroup"
        ):
            if field_obj.get(k) is not None:
                val = field_obj[k]
                break

    if val is None and isinstance(field_obj.get("valueArray"), list):
        parts = [str(best_value(v)) for v in field_obj["valueArray"]]
        val = " | ".join(parts) if parts else None

    if val is None and isinstance(field_obj.get("valueObject"), dict):
        parts = [f"{k}: {best_value(v)}" for k, v in field_obj["valueObject"].items()]
        val = " | ".join(parts) if parts else None

    return val


def extract_from_document_fields(data, field_names):
    extracted = {}
    docs = data.get("analyzeResult", {}).get("documents", [])

    for doc in docs:
        fields = doc.get("fields", {})
        if not isinstance(fields, dict):
            continue

        for name in field_names:
            if name in fields:
                field_obj = fields[name]
                value = best_value(field_obj)
                conf = field_obj.get("confidence")
                conf_pct = round(conf * 100, 2) if isinstance(conf, (int, float)) else None

                if name not in extracted:
                    extracted[name] = value
                    extracted[f"{name}_confidence"] = conf_pct

    return extracted


def make_file_id(stem: str) -> str:
    for suffix in ("_result", "_output", "_analysis", "_analyze"):
        if stem.endswith(suffix):
            return stem[:-len(suffix)]
    return stem


# ============================================================
# === TABLE STRING PARSING (EXACT CONTENT, CORRECT ROWS)     ===
# ============================================================
ROW_HEADER_RE = re.compile(r'(?m)(^|\n|\|)\s*(\d+)\s*:')  # numeric rows like "2:"
KV_RE = re.compile(r'^\s*([A-Za-z0-9_]+)\s*:\s*(.*?)\s*$')  # key/value like "s1: None"
SECTION_KEY_RE = re.compile(r'^[A-Z]{2}$')

# Keys that are legitimate “data keys” inside a numeric row
LEGIT_ROW_KEYS = {
    "time", "Response", "Verbatim", "verbatim",
    "r1", "r2", "r3", "r4", "r5",
    "dk", "nr",
    "s0", "s1", "s2",
}

def looks_like_table_string(s):
    return isinstance(s, str) and ROW_HEADER_RE.search(s)

INLINE_SECTION_RE = re.compile(
    r'(^|\|)\s*([A-Za-z0-9_]{1,20})\s*:\s*[A-Za-z0-9_]+\s*:'
)

def looks_like_inline_section(s):
    return isinstance(s, str) and INLINE_SECTION_RE.search(s)


def normalise_key(prefix: str, key: str) -> str:
    """
    Ensure key naming matches the required Excel column names.
    - For si_p*, vc_p*, co_p*, in_p12: force 'Verbatim' -> 'verbatim'
    - Leave in_p13 'Response' untouched (your order list uses 'Response' there)
    """
    if not isinstance(key, str):
        return key

    # Force lower-case verbatim for these prefixes
    if prefix.startswith(("si_p", "vc_p", "co_p", "in_p12")) and key.lower() == "verbatim":
        return "verbatim"

    # Otherwise leave as-is (e.g., Response for in_p13)
    return key


def parse_table_string(prefix, s):
    """
    Expands flattened 'table strings' into separate columns, WITHOUT altering values.

    Critical fix:
    If a section marker like 'SB:' appears inside the chunk for row '2',
    we stop parsing that chunk immediately so row 2 doesn't absorb SB's selection.
    """
    out = {}
    if not isinstance(s, str) or not s.strip():
        return out

    s = s.replace("|", "\n")

    rows = list(ROW_HEADER_RE.finditer(s))
    for i, m in enumerate(rows):
        row = m.group(2)
        start = m.end()
        end = rows[i + 1].start() if i + 1 < len(rows) else len(s)
        chunk = s[start:end].strip()

        current_sub = None
        for line in chunk.splitlines():
            line = line.strip()
            if not line:
                continue

            kv = KV_RE.match(line)
            if not kv:
                continue

            key, val = kv.groups()
            key = key.strip()
            val = val.strip()

            # Normalise keys to match desired Excel column names
            key = normalise_key(prefix, key)

            # --- sr_p19 special handling (DA/DB/DC nested blocks) ---
            if prefix == "sr_p19" and row == "12" and key in ("DA", "DB", "DC"):
                current_sub = key
                continue

            if prefix == "sr_p19" and row == "12" and current_sub:
                out[f"{prefix}_{row}_{current_sub}_{key}"] = val
                continue
            # -------------------------------------------------------

            # --- stop when a section marker appears ---
            if SECTION_KEY_RE.match(key) and key not in LEGIT_ROW_KEYS:
                break

            out[f"{prefix}_{row}_{key}"] = val

    return out

###new

def parse_inline_sections(prefix, s):
    """
    Parses inline sections like:
    | Sample: Verbatim: X | dk: None | ...
    | DA: r4: :selected: | r1: None | ...
    | d: r1: :selected: | ... | s: r1: :selected: | ...
    | ra: verbatim: 8126 | s: None | rb: verbatim: None | s: None
    | sa: verbatim: X | dk: None | ...
    | 1_1: verbatim: 57 | t1_2: :selected: | dk: None | ...
    | sa: 1: verbatim: X | dk: None | ...
    | sa_1: verbatim: X | dk: None | ...

    Produces:
    ds_sa_1_verbatim
    lns_sb_2_t1_2
    db_p14_sa_1_dk
    db_p15_10_2_verbatim
    etc.
    """
    out = {}
    if not isinstance(s, str) or not s.strip():
        return out

    parts = [p.strip() for p in s.split("|") if p.strip()]

    current_section = None

    # Section labels that are always true sections
    SECTION_LABELS = {
        "DA", "DB", "DC",
        "SA", "SB",
        "Sample",
        "d",
        "ra", "rb",
        "sa", "sb"
    }

    # Lower-case "s" is ambiguous:
    # - in vp_l and ar_l it is a section
    # - in rd it is a key inside ra/rb
    ALLOW_S_SECTION_PREFIXES = {"vp_l", "ar_l"}

    # Fields that use trial labels like 1_1, 1_2, 10_1, etc.
    NUMERIC_TRIAL_SECTION_PREFIXES = {
        "df", "ds", "lns",
        "db", "db_p14", "db_p15"
    }

    # Fields that may use nested forms like:
    # sa: 1: verbatim: X
    # sb: 2: verbatim: X
    NESTED_SAMPLE_SECTION_PREFIXES = {
        "ds", "lns",
        "db", "db_p14", "db_p15"
    }

    def is_numeric_trial_label(label):
        return bool(re.match(r"^\d+_\d+$", label))

    def is_sample_trial_label(label):
        return bool(re.match(r"^(sa|sb|s)_\d+$", label, flags=re.IGNORECASE))

    def is_section_label(possible_section):
        """
        Decide whether the token before the first colon is a section label
        or just a key within the current section.
        """
        if possible_section in SECTION_LABELS:
            return True

        if possible_section == "s" and prefix in ALLOW_S_SECTION_PREFIXES:
            return True

        if prefix in NUMERIC_TRIAL_SECTION_PREFIXES and is_numeric_trial_label(possible_section):
            return True

        if prefix in NESTED_SAMPLE_SECTION_PREFIXES and is_sample_trial_label(possible_section):
            return True

        return False

    for part in parts:

        # Keep consuming leading section labels if the chunk starts with nested sections.
        #
        # Examples:
        # "1_1: verbatim: 57"
        #   -> section = "1_1", remaining = "verbatim: 57"
        #
        # "sa_1: verbatim: X"
        #   -> section = "sa_1", remaining = "verbatim: X"
        #
        # "sa: 1: verbatim: X"
        #   -> section = "sa_1", remaining = "verbatim: X"
        while ":" in part:
            first_split = part.split(":", 1)
            possible_section = first_split[0].strip()
            remainder = first_split[1].strip()

            # Standard section labels
            if is_section_label(possible_section):
                current_section = possible_section
                part = remainder

                if not part:
                    break

                continue

            # Nested numeric under sa/sb:
            # sa: 1: verbatim: X -> sa_1_verbatim
            # sb: 2: verbatim: X -> sb_2_verbatim
            if (
                prefix in NESTED_SAMPLE_SECTION_PREFIXES
                and current_section in {"sa", "sb"}
                and re.match(r"^\d+$", possible_section)
            ):
                current_section = f"{current_section}_{possible_section}"
                part = remainder

                if not part:
                    break

                continue

            # If it is not a section, stop and treat it as key:value
            break

        # Process key:value inside current section
        if ":" in part and current_section is not None:
            key, val = [x.strip() for x in part.split(":", 1)]

            key = normalise_key(prefix, key)

            if ":selected:" in val:
                out[f"{prefix}_{current_section}_{key}"] = ":selected:"
            else:
                out[f"{prefix}_{current_section}_{key}"] = val

    return out

# ============================================================
# === MAIN                                                  ===
# ============================================================
def resolve_input_files(path: Path):
    if path.is_file():
        if path.suffix.lower() != ".json":
            raise ValueError(f"Input file must be a .json file: {path}")
        return [path]

    if path.is_dir():
        return sorted(path.glob("*.json"))

    raise FileNotFoundError(f"Input path not found: {path}")


def main():
    files = resolve_input_files(input_path)
    if not files:
        raise FileNotFoundError(f"No JSON files found in {input_path}")

    rows = []

    for fp in files:
        with open(fp, encoding="utf-8") as f:
            data = json.load(f)

        extracted = extract_from_document_fields(data, fields_to_extract)

        # Note: "SourceFile" removed to match your desired final schema (it wasn't in your order list)
        row = {
            "FileID": make_file_id(fp.stem),
        }

        for f in fields_to_extract:
            v = extracted.get(f)

            # ALWAYS-Raw fields -> store as <field>_raw only
            if f in ALWAYS_RAW_FIELDS:
                row[f"{f}_raw"] = v

            # Table-like fields -> keep raw + expand (without altering values)
            elif looks_like_table_string(v):
                row[f"{f}_raw"] = v

                # existing row-based parsing
                row.update(parse_table_string(f, v))

                # inline section parsing
                row.update(parse_inline_sections(f, v))

            # Inline-section fields without numeric row headers (df, ds, db)
            elif looks_like_inline_section(v):
                row[f"{f}_raw"] = v
                row.update(parse_inline_sections(f, v))

            # Normal scalar fields
            else:
                row[f] = v

            # Confidence always appended
            row[f"{f}_confidence"] = extracted.get(f"{f}_confidence")

        rows.append(row)

    df = pd.DataFrame(rows)

    # Enforce the exact final Excel column order
    df = apply_final_column_order(df)

    output_excel_path.parent.mkdir(parents=True, exist_ok=True)
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    out = output_excel_path.with_name(f"{output_excel_path.stem}_{ts}.xlsx")

    df.to_excel(out, index=False, engine="openpyxl")
    print(f"✅ Saved {len(df)} rows to {out}")


if __name__ == "__main__":
    try:
        main()
    except Exception:
        print(traceback.format_exc())
