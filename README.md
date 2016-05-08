# xls2lua
将excel导出成lua，游戏配置导出使用

### NOTICE
初步实现，BUG很多，代码很烂，欢迎帮忙改进

需要xlrd

根据字段颜色区分导出标记[客户端，服务端，多语言],支持多张子表合并导出,一个excel为一个导出，空数据导出为nil，非config开头的xls文件不转换

lua文件使用module，所以lua5.2以上的请使用__tools.lua 中提供的module实现，或者修改导出格式

### 使用
main.py [file]
提供两个导出bat
export.cmd 创建快捷方式后放在配置表所在目录，拖配置表到快捷方式上单个导出
export_all.cmd 导出所有配置表

导出后
```lua
-- client side
require("defineExample")
require("configExample")
local name = resmng.configExample[resmng.ITEM_1].Name --auto change to specific language by using meta

-- server side
require("defineExample")
require("configExample")
local name = resmng.configExample[resmng.ITEM_1].ID
```
### 导出格式说明
导出分为define文件和config文件,define为配置表数值key值，config为数据
配置表第一二列格式如下时导出define,其他情况不导出
```
ID      EnumID
ITEM_1  1
ITEM_2  2
```
时导出为
```
ITEM_1 = 1
ITEM_2 = 2
```

export同级目录生成_out文件夹
```
_out
  |-client
    |-defineLanguage
    |-configLanguageCN
    |-configLanguageEN
    |-...
  |-server
    |-defineLanguage
    |-...
```
configLanguage为必须的配置表，提供多语言支持，
多语言配置表包括defineLanguage 和 configLanguage, configLanguage按照语言分成多个文件根据语言加载不同的配置表

```lua
-- defineLanguage
module("resmng")

LG_ITEM_1=1
LG_ITEM_2=2
```
```lua
-- defineLanguageCN
module("resmng")

configLanguage = {
    LG_ITEM_1="物品1",
    LG_ITEM_2="物品2",
}
```
客户端导出的config标记为多语言字段的将使用Language表实现多语言翻译
基本思路为使用meta自动转换
```
module("resmng")
-- 数据表
configExample = {
    [ITEM_1]={ITEM_1,LG_ITEM_1,"icon1",100,1,0,},
    [ITEM_1]={ITEM_1,LG_ITEM_2,"icon2",200,2,1,},
}
-- 数据表字段名
_configExampleKey = {
    ID = 1,
    Name = 2,
    Icon = 3,
    Price = 4,
}
-- 多语言字段
_configExampleLanguageKey = {
    Name = 1,
}
-- 加载表时对每行数据设置meta，方便使用
_mtExample = {
	__index = function(t, k)
		local a = t[_configExampleKey[k]]
		if _configExampleLanguageKey[k] ~= nil then
			return configLanguage[t[_configExampleKey[k]]]
		else
			return t[_configExampleKey[k]]
		end
	end
}

for k, v in pairs(configExample) do
	setmetatable(v, _mtExample)
end

```
