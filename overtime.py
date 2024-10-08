import pandas as pd


class OvertimeStastics:
    """Overtime Stastics"""

    ADI_SORT_KEY = {
        # units
        "署長室": 1,
        "陳副署長室": 2,
        "林副署長室": 3,
        "主任秘書室": 4,
        "政策規劃組": 5,
        "前瞻政策科": 6,
        "產業人才科": 7,
        "計畫管理科": 8,
        "綜合業務科": 9,
        "通訊傳播組": 10,
        "基礎環境科": 11,
        "傳播推廣科": 12,
        "通訊應用科": 13,
        "平臺經濟組": 14,
        "平臺應用科": 15,
        "平臺治理科": 16,
        "數位體驗科": 17,
        "數據經濟科": 18,
        "新興跨域組": 19,
        "資安產業科": 20,
        "全球運籌科": 21,
        "場域實證科": 22,
        "地方鏈結科": 23,
        "數位服務組": 24,
        "軟體產業科": 25,
        "轉型輔導科": 26,
        "整合服務科": 27,
        "創新應用科": 28,
        "秘書室": 29,
        "文書科": 30,
        "事務科": 31,
        "人事室": 32,
        "政風室": 33,
        "主計室": 34,
        # jobs
        "署長": 1,
        "副署長": 2,
        "主任秘書": 3,
        "組長": 4,
        "主任": 5,
        "副組長": 6,
        "簡任技正": 7,
        "簡任視察": 8,
        "專門委員": 9,
        "科長": 10,
        "技正": 11,
        "視察": 12,
        "專員": 13,
        "科員": 14,
        "助理員": 15,
        "專案規劃師": 16,
        "專案分析師": 17,
        "資安系統分析師": 18,
    }
    ACS_SORT_KEY = {
        "署長室": 0,
        "林副署長室": 1,
        "鄭副署長室": 2,
        "主任秘書室": 3,
        "綜合規劃組": 4,
        "策略規劃科": 5,
        "計畫審查科": 6,
        "組織人力科": 7,
        "綜合業務科": 8,
        "應用服務科": 9,
        "通報應變組": 10,
        "設施防護科": 11,
        "應變處理科": 12,
        "威脅分析科": 13,
        "情勢研析科": 14,
        "輔導培訓組": 15,
        "基準發展科": 16,
        "人才培訓科": 17,
        "機關輔導科": 18,
        "教育推廣科": 19,
        "稽核檢查組": 20,
        "稽核一科": 21,
        "稽核二科": 22,
        "稽核及分析科": 23,
        "檢測演練科": 24,
        "法規及國合組": 25,
        "法規推動科": 26,
        "法制行政科": 27,
        "國際合作科": 28,
        "秘書室": 29,
        "秘書室一科": 30,
        "秘書室二科": 31,
        "人事室": 32,
        "政風室": 33,
        "主計室": 34,
        "署長": 35,
        "副署長": 36,
        "主任秘書": 37,
        "組長": 38,
        "副組長": 39,
        "高級分析師": 40,
        "簡任視察": 41,
        "主任": 42,
        "科長": 43,
        "代理科長": 44,
        "分析師": 45,
        "設計師": 46,
        "助理程式設計師": 47,
        "視察": 48,
        "專員": 49,
        "科員": 50,
        "助理員": 51,
        "書記": 52,
        "高級資安程式設計師": 53,
        "資安系統分析師": 54,
        "資安程式設計師": 55,
        "專案分析師": 56,
    }

    CH_COlS = [
        "單位名稱",
        "職稱",
        "姓名",
        "加班日期",
        "核可時數",
        "已請款時數",
        "已補休時數",
        "剩餘可用時數",
    ]
    EN_COLS = [
        "unit",
        "job",
        "name",
        "year-month",
        "appr",
        "paid",
        "rest",
        "remain",
    ]
    CAL_COLS = ["appr", "paid", "rest", "remain"]
    CAL_COLS_MAP = {
        "appr-H": "核可(時)",
        "appr-M": "核可(分)",
        "paid-H": "請款(時)",
        "paid-M": "請款(分)",
        "rest-H": "補休(時)",
        "rest-M": "補休(分)",
        "remain-H": "剩餘(時)",
        "remain-M": "剩餘(分)",
    }
    NON_CAL_COLS = ["unit", "job", "name", "year-month"]

    def __init__(self, file: str, sort_key: str = "adi"):

        # determine sort key
        if sort_key == "adi":
            self.sort_key = self.ADI_SORT_KEY
        elif sort_key == "acs":
            self.sort_key = self.ACS_SORT_KEY
        else:
            raise ValueError("Invalid sort key")

        self.source_df = self._clean(pd.read_excel(file))

    def _clean(self, df: pd.DataFrame):
        """Clean DataFrame"""

        # 1) reformat DataFrame
        column_row: int = df[df.iloc[:, 0] == "單位名稱"].index.values[0]
        start_row: int = column_row + 1
        df.columns = df.iloc[column_row, :]
        df = df.iloc[start_row:][self.CH_COlS].copy()
        df.rename(columns=dict(zip(self.CH_COlS, self.EN_COLS)), inplace=True)
        df.columns.name = None
        df["year-month"] = (
            df["year-month"]
            .str.split("/")
            .map(lambda x: str(int(x[0])) + "-" + str(int(x[1])))
        )

        # 2) split hours and minutes
        df.fillna("0-0", inplace=True)
        for col in self.CAL_COLS:
            df[f"{col}-H"] = df[col].str.split("-").map(lambda x: x[0])
            df[f"{col}-M"] = df[col].str.split("-").map(lambda x: x[1])

        df.drop(columns=self.CAL_COLS, inplace=True)

        # 3) sort DataFrame
        df = df.sort_values(
            by=["unit", "job"], key=lambda x: x.map(self.sort_key)
        )
        df.reset_index(drop=True, inplace=True)

        return df

    def execute(self, file: str = "output.xlsx"):
        """Execute procedure"""

        personal_data = self._personal_data()
        result = pd.DataFrame(personal_data).T
        result.fillna(0, inplace=True)
        result = self._round_minutes_up(result).copy()
        result.reset_index(names=["name"], inplace=True)
        result = result.set_axis(range(1, len(result) + 1), axis="index")
        result = self._reformat_columns(result)
        self._export(result, file)

        return result

    def _personal_data(self) -> dict:
        """
        Generate personal data dictionary
        :return: personal data dictionary -> {
            name: {[year-month]appr_H: time, [year-month]appr_M: time, ...},
            ...
        }
        """

        result = {}
        names = self.source_df["name"].unique()
        for name in names:
            name_df = self.source_df[self.source_df["name"] == name].copy()
            name_df.reset_index(drop=True, inplace=True)
            # assign personal basic info (unit and job)
            result[name] = {
                k: v
                for k, v in zip(
                    ["unit", "job"],
                    [name_df.loc[0, "unit"], name_df.loc[0, "job"]],
                )
            }

            result[name].update(self._sumup_year_month_time(name_df))

        return result

    def _sumup_year_month_time(self, df) -> dict:
        """★ Sum up personal year-month time ★"""

        result = {}
        cal_cols = [col for col in df.columns if col not in self.NON_CAL_COLS]
        year_month_list = df["year-month"].unique()
        for y_m in year_month_list:
            y_m_df = df[df["year-month"] == y_m]
            for col in cal_cols:
                year, month = y_m.split("-")[0], y_m.split("-")[1]
                if int(month) < 10:
                    month = "0" + month
                # convert datatype to int
                y_m_df.loc[:, col] = y_m_df.loc[:, col].astype(int)
                result = {
                    **result,
                    **{f"[{year}-{month}]_{col}": y_m_df[col].sum()},
                }

        return result

    def _round_minutes_up(self, df) -> pd.DataFrame:
        """Round minutes up to hours"""

        result = df.copy()
        cal_cols = [col for col in df.columns if col not in self.NON_CAL_COLS]
        year_month_list = set([col.split("_")[0] for col in cal_cols])
        for y_m in year_month_list:
            for col in self.CAL_COLS:
                hour_col, min_col = f"{y_m}_{col}-H", f"{y_m}_{col}-M"
                result[hour_col] = result[hour_col] + (
                    result[min_col] / 60
                ).astype(int)
                result[min_col] = (result[min_col] % 60).astype(int)

        return result

    def _reformat_columns(self, df: pd.DataFrame):
        """Reformat columns"""

        # sort columns
        basicinfo_cols = ["unit", "job", "name"]
        cal_cols = sorted(
            [col for col in df.columns if col not in basicinfo_cols]
        )
        result = df[basicinfo_cols + cal_cols]

        # rename columns
        # basic info
        result.rename(
            columns={"unit": "單位", "job": "職稱", "name": "姓名"},
            inplace=True,
        )
        # calulation info
        for col in cal_cols:
            y_m, cal_attr = col.split("_")[0], col.split("_")[1]
            result.rename(
                columns={col: f"{y_m}\n{self.CAL_COLS_MAP[cal_attr]}"},
                inplace=True,
            )

        return result

    def _export(self, df: pd.DataFrame, file: str):
        """Export DataFrame to Excel file"""

        sheets = ["原始資料", "核可", "請款", "補休", "剩餘"]
        writer = pd.ExcelWriter(file, engine="xlsxwriter")
        wb = writer.book

        for sheet in sheets:
            if sheet == "原始資料":
                export_df = df.copy()
            else:
                cols = [col for col in df.columns if sheet in col]
                export_df = df[list(df.columns[:3]) + cols].copy()

            export_df.to_excel(writer, sheet_name=sheet, index=False)

            format = wb.add_format(
                {
                    "align": "center",
                    "valign": "vcenter",
                    "text_wrap": True,
                    "font_name": "微軟正黑體",
                    "font_size": 12,
                }
            )

            sheet = writer.sheets[sheet]
            for col_num, value in enumerate(export_df.columns.values):
                sheet.write(0, col_num, value, format)

            sheet.set_column(0, 2, 13, format)
            sheet.set_column(3, len(export_df.columns), 10, format)

        writer.close()
