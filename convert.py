# coding:utf-8

import argparse
import decimal
import datetime
import re
import openpyxl

標準税率 = decimal.Decimal("0.1") #10%
軽減税率 = decimal.Decimal("0.08") #8%

丸めモード = decimal.ROUND_HALF_UP #四捨五入

品目種別 = ["商品", "送料"] #値引きとかも必要かも

assert type(標準税率) == decimal.Decimal
assert type(軽減税率) == decimal.Decimal

class 品目cls:
    """請求書に記載する品目"""
    def __init__(self, 種別, 軽減税率flg, 品名, 単価, 個数): #小計は計算する
        assert 種別 in 品目種別
        assert type(軽減税率flg) == bool
        assert type(単価) == int or type(単価) == decimal.Decimal
        assert type(個数) == int or type(個数) == decimal.Decimal
        self.種別 = 種別
        self.軽減税率flg = 軽減税率flg
        self.品名 = 品名
        self.単価 = 単価
        self.個数 = 個数
        self.update_小計()
    
    def update_小計(self):
        self.小計 = self.単価 * self.個数
    
    def __str__(self):
        return "品目({}, {}, {}, {}円, {}個)".format(
                    self.種別,
                    "軽減" if self.軽減税率flg else "標準",
                    self.品名,
                    self.単価,
                    self.個数 )
    
    def __lt__(self, other): #ソート用
        def ひらがなをカタカナに(text):
            delta = ord("ァ") - ord("ぁ")
            trans_dic = {chr(n): chr(n+delta) for n in range(ord("ぁ"), ord("ゖ"))}
            return text.translate(str.maketrans(trans_dic))
            
        #送料は後に表示する
        if 品目種別.index(self.種別) != 品目種別.index(other.種別):
            return 品目種別.index(self.種別) < 品目種別.index(other.種別)
        
        # 100サイズ, 120サイズ, 80サイズの順に並ぶのは気持ち悪い
        #品名が数字を除いて同じ場合は、数字の数値の大小でソートするようにできればよさそう
        ptn = re.compile(r"(\D*)(\d+)")
        カタカナ品名 = ひらがなをカタカナに(self.品名)
        カタカナ品名_other = ひらがなをカタカナに(other.品名)
        m = ptn.findall(カタカナ品名)
        n = ptn.findall(カタカナ品名_other)
        if m and n: #両方が数字を含む
            for item, other_item in zip(m, n):
                text = item[0]
                number = int(item[1])
                other_text = other_item[0]
                other_number = int(other_item[1])
                if text != other_text:
                    return text < other_text
                if number != other_number:
                    return number < other_number
            #これで判断がつかなければそのままの品名の比較でいいや↓
        
        #仮名をカタカナに統一して品名でソート
        if カタカナ品名 != カタカナ品名_other:
            return カタカナ品名 < カタカナ品名_other
        #品名でソート
        return self.品名 < other.品名
        
        
        
    def __eq__(self, other): #in演算子用
        if type(other) != type(self):
            return False
        if       self.種別 == other.種別 \
             and self.品名 == other.品名 \
             and self.軽減税率flg == other.軽減税率flg \
             and self.単価 == other.単価:
                return True
        return False
    
    def add_num(self, num): #num個追加
        assert type(num) == int
        self.個数 += num
        self.update_小計()
    
    def get品名(self):
        return self.品名 + ("（※）" if self.軽減税率flg else "")
    

class 税込請求書cls:
    """出力のための請求書クラス。税込金額で記載"""
    def __init__(self, ID, 請求先, 年月日, 品目lst):
        #合計, 標準税率対象額, 標準税額, 軽減税率対象額, 軽減税額, 消費税計は計算する
        self.ID = ID
        self.請求先 = 請求先
        assert type(年月日) == datetime.datetime or type(年月日) == datetime.date
        self.年月日 = 年月日
        self.品目lst = 品目lst
        self.update_合計と税()

    def update_合計と税(self):
        合計 = 0
        標準税率対象額 = 0
        軽減税率対象額 = 0
        for 品目 in self.品目lst:
            合計 += 品目.小計
            if 品目.軽減税率flg:
                軽減税率対象額 += 品目.小計
            else:
                標準税率対象額 += 品目.小計
        #decimal.Decimal.quantize()を利用して整数に
        標準税額 = (標準税率対象額 - 標準税率対象額/(1+標準税率)).quantize(decimal.Decimal("0"), rounding=丸めモード)
        軽減税額 = (軽減税率対象額 - 軽減税率対象額/(1+軽減税率)).quantize(decimal.Decimal("0"), rounding=丸めモード)
        消費税計 = 標準税額 + 軽減税額
        
        self.合計金額 = 合計
        self.標準税率対象額 = 標準税率対象額
        self.標準税額 = 標準税額
        self.軽減税率対象額 = 軽減税率対象額
        self.軽減税額 = 軽減税額
        self.消費税計 = 消費税計
        
    def __str__(self):
        return "税込請求書(ID: {}, {}, {}, 合計{}, 標準対象{}, 標準税{}, 軽減対象{}, 軽減税{}, 税計{}, [{}])".format(
                        self.ID,
                        self.請求先,
                        self.年月日.isoformat(),
                        self.合計金額,
                        self.標準税率対象額,
                        self.標準税額,
                        self.軽減税率対象額,
                        self.軽減税額,
                        self.消費税計,
                        ", ".join(map(str, self.品目lst)) )
    
    def add_品目(self, 品目):
        if 品目 in self.品目lst: #すでに存在する場合
            self.品目lst[self.品目lst.index(品目)].add_num(品目.個数)
        else: #まだ存在しない
            self.品目lst.append(品目)
        self.update_合計と税()
    
    def add_品目lst(self, 品目lst):
        for 品目 in 品目lst:
            self.add_品目(品目)
        
                        
class Excel出力:
    def __init__(self):
        self.wb = openpyxl.Workbook()
        self.ws = self.wb.active
        self.ws.title = "請求書"
        self.ws.append([    "請求書ID",
                            "年",
                            "月",
                            "日",
                            "被請求者",
                            "品目",
                            "単価",
                            "個数",
                            "品目小計",
                            "代金計",
                            "標準税率対象",
                            "軽減税率対象",
                            "標準税",
                            "軽減税",
                            "消費税計" ])
    
    def output_税込請求書lst(self, 税込請求書lst):
        for 税込請求書 in 税込請求書lst:
            self.output_税込請求書(税込請求書)
        
    def output_税込請求書(self, 税込請求書):
        請求書ID = 税込請求書.ID
        年月日 = 税込請求書.年月日
        年 = 年月日.year
        月 = 年月日.month
        日 = 年月日.day
        被請求者 = 税込請求書.請求先
        
        代金計 = 税込請求書.合計金額
        標準税率対象額 = 税込請求書.標準税率対象額
        軽減税率対象額 = 税込請求書.軽減税率対象額
        標準税額 = 税込請求書.標準税額
        軽減税額 = 税込請求書.軽減税額
        消費税計 = 税込請求書.消費税計
        
        for item in sorted(税込請求書.品目lst):
            #if item.種別 == "商品":
            品目 = item.get品名()
            単価 = item.単価
            個数 = item.個数
            品目小計 = item.小計
            row = [請求書ID, 年, 月, 日, 被請求者, 品目, 単価, 個数, 品目小計, 代金計, 標準税率対象額, 軽減税率対象額, 標準税額, 軽減税額, 消費税計]
            self.ws.append(row)
            # elif item.種別 == "送料"
            # 種別 = "送料"
            # 品目 = "送料 " + "・".join(x.get送り先地域()) + " " + x.重さ
            # 単価 = x.送料単価
            # 個数 = x.個数
            # 品目小計 = 単価*個数
            # row = [請求書ID, 年, 月, 日, 種別, 被請求者, 品目, 単価, 個数, 品目小計, 代金計, 標準税率対象料金, 軽減税率対象料金, 標準税, 軽減税, 消費税計]
            # self.ws.append(row)
        self.ws.append([None, None, None, None, None, None, None, None, None, 代金計, 標準税率対象額, 軽減税率対象額, 標準税額, 軽減税額, 消費税計])

    def save(self, filename):
        self.wb.save(filename)

class Excel変換器:
    def __init__(self, excelfile):
        self.wb = openpyxl.load_workbook(excelfile, data_only=True) #数式ではなく値を読み込む
        self.load送料表()
        
    def load送料表(self):
        ws = self.wb["都道府県送料"]
        #送料表[都道府県][サイズ] -> 料金, 送付先地域名称
        #とする。サイズごとに同一料金をグルーピングして、地域名の連結で名称とする
        vals = ws.values
        heading = next(vals)
        サイズlst = heading[2:]
        
        送料表 = {}
        名称表 = {}
        
        グループ = {}
        都道府県lst = []
        for row in vals:
            都道府県 = row[0]
            地域 = row[1]
            for サイズ in サイズlst:
                idx = heading.index(サイズ)
                送料 = row[idx]

                if (サイズ, 送料) not in グループ:
                    グループ[(サイズ, 送料)] = [地域]
                else:
                    if 地域 not in グループ[(サイズ, 送料)]:
                        グループ[(サイズ, 送料)].append(地域)
                送料表[(都道府県, サイズ)] = 送料
            都道府県lst.append(都道府県)
        for 都道府県 in 都道府県lst:
            for サイズ in サイズlst:
                送料 = 送料表[(都道府県, サイズ)]
                名称 = "・".join(グループ[(サイズ, 送料)])
                名称表[(都道府県, サイズ)] = 名称+" "+str(サイズ)+"サイズ"
        self.送料表 = 送料表 #使ってない
        self.名称表 = 名称表
    
    def get税込送料品目(self, 都道府県, サイズ, 単価, 個数):
        #placeholder
        品名 = "送料 " + self.名称表[(都道府県, サイズ)]
        return 品目cls("送料", False, 品名, 単価, 個数)

    def get_品目lst_from_row(self, row):
        #請求書ID	日付	被請求者	商品品目	軽減	単価	個数	送り先id	送り先〒	送り先住所	サイズ	送料単価	送料個数
        請求書ID = row[0]
        # 日付 = row[1]
        # 被請求者 = row[2]
        商品品目 = row[3]
        軽減対象 = True if row[4] else False #空でなければTrue
        商品単価 = row[5]
        商品個数 = row[6]
        #送り先id = row[7]
        #送り先郵便番号 = row[8]
        送り先都道府県 = row[9]
        送料サイズ = row[10]
        送料単価 = row[11]
        送料個数 = row[12]
        
        assert 請求書ID is not None
        品目lst = []
        if 商品品目: 
            品目lst.append( 品目cls("商品", 軽減対象, 商品品目, 商品単価, 商品個数) )
        if 送り先都道府県:
            if 送料個数 is not None and 送料個数>0:
                品目lst.append( self.get税込送料品目(送り先都道府県, 送料サイズ, 送料単価, 送料個数) )
            
        return 品目lst

    def get_税込請求書_from_row(self, row):
        #請求書ID	日付	被請求者	商品品目	軽減	単価	個数	送り先id	送り先〒	送り先住所	サイズ	送料単価	送料個数
        請求書ID = row[0]
        日付 = row[1]
        被請求者 = row[2]
        
        assert 請求書ID is not None
        品目lst = self.get_品目lst_from_row(row)
            
        return 税込請求書cls(請求書ID, 被請求者, 日付, 品目lst)

    def convert(self, dest_filename, min_id=None, max_id=None, date=None):
        ws = self.wb.worksheets[0] #最初のシート
        vals = ws.values
        _ = next(vals) #タイトル行をスキップ
        請求書dict = {}
        for row in vals:
            if [e for e in row if e is not None]: #すべてNoneの場合を除く
                請求書ID = row[0]
                if min_id is not None and 請求書ID < min_id: #min_idより請求書IDが小さい場合は出力しない
                    continue
                if max_id is not None and 請求書ID > max_id: #max_idより請求書IDが大きい場合は出力しない
                    continue
                日付 = row[1] #日付が指定してある場合、日付が異なれば無視
                if date is not None and (日付.year != date.year or 日付.month != date.month or 日付.day != date.day):
                    continue
                    
                if 請求書ID not in 請求書dict:
                    請求書dict[請求書ID] = self.get_税込請求書_from_row(row)
                else:
                    請求書dict[請求書ID].add_品目lst(self.get_品目lst_from_row(row))
        
        excel_outputter = Excel出力()
        excel_outputter.output_税込請求書lst(請求書dict.values())
        excel_outputter.save(dest_filename)

def main(excelfile, min_id, max_id, date):
    excel_translator = Excel変換器(excelfile)
    excel_translator.convert("processed.xlsx", min_id, max_id, date)
    
    # 請求書a = 税込請求書cls(99, "愛城華恋", datetime.date.fromisoformat("2022-01-06"), [
        # 品目cls("商品", True, "トマト 2kg箱", 3000, 1),
        # 品目cls("商品", True, "りんご 10kg箱", 7000, 2),
        # 品目cls("送料", False, "送料 東北・関東・信越・北陸・東海 80サイズ", 1100, 1),
        # 品目cls("送料", False, "送料 東北・関東・信越・北陸・東海 120サイズ", 1590, 2)] )
    # 請求書b = 税込請求書cls(100, "神楽ひかり", datetime.date.fromisoformat("2022-01-06"), [
        # 品目cls("商品", True, "りんご 5kg箱", 3500, 2),
        # 品目cls("商品", True, "バナナ 5kg箱", 2500, 1),
        # 品目cls("送料", False, "送料 東北・関東・信越・北陸・東海 120サイズ", 1590, 1)] )
    # 請求書c = 税込請求書cls(101, "星見純那", datetime.date.fromisoformat("2022-01-06"), [
        # 品目cls("商品", True, "ぶどう 2kg箱", 6500, 1),
        # 品目cls("商品", True, "バナナ 5kg箱", 2500, 1),
        # 品目cls("送料", False, "送料 東北・関東・信越・北陸・東海 80サイズ", 1100, 1),
        # 品目cls("送料", False, "送料 東北・関東・信越・北陸・東海 100サイズ", 1330, 1)] )
    
    # excel_outputter = Excel出力()
    # excel_outputter.output_税込請求書lst([請求書a, 請求書b, 請求書c])
    # excel_outputter.save("processed.xlsx")


if __name__ == '__main__':
    def str_to_date(text):
        if not text:
            return None
        try:
            return datetime.date.fromisoformat(text)
        except ValueError:
            n = re.match(r"^(\d{4})/([012]?\d)/([0123]?\d)$", text)
            if n:
                return datetime.date(year=int(n[1]), month=int(n[2]), day=int(n[3]))
            else:
                raise ValueError("日付指定の書式が異常です")
    
    parser = argparse.ArgumentParser(description='請求書の書式を変換します')
    parser.add_argument('filename', help='変換元ファイル名(excel file)')
    parser.add_argument('--min_id', type=int, help='このID以上の請求書を出力')
    parser.add_argument('--max_id', type=int, help='このID以下の請求書を出力')
    parser.add_argument('--date', type=str_to_date, help='日付を指定して出力(yyyy-MM-dd)')
    
    args = parser.parse_args()

    main(args.filename, args.min_id, args.max_id, args.date)
    
