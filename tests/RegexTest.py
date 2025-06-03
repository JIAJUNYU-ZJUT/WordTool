import re

sql1 = """CREATE TABLE IF NOT EXISTS dw_jx.`ods_glorious_mission_task_ds_bak20220906_orc`(
`id` STRING COMMENT '',
`business_id` STRING COMMENT '业务id',
`org_id` STRING COMMENT '集团',
`task_type_code` STRING COMMENT '任务类型code',
`button_code` STRING COMMENT '任务按钮code',
`executor` STRING COMMENT '执行人',
`operator` STRING COMMENT '操作人',
`shop_code` STRING COMMENT '所在店铺id',
`task_status` INT COMMENT '任务状态0:开始,1:已完成，2:无效',
`field_related_content` STRING COMMENT '关联内容(json)',
`begin_date` BIGINT COMMENT '任务开始时间',
`end_date` BIGINT COMMENT '任务结束时间',
`deal_date` BIGINT COMMENT '任务处理时间',
`create_date` BIGINT COMMENT '任务创建时间戳(毫秒级)',
`date_create` TIMESTAMP COMMENT '',
`date_update` TIMESTAMP COMMENT '',
`date_delete` TIMESTAMP COMMENT '')
 COMMENT '任务表'
 PARTITIONED BY (
`ds` STRING COMMENT '分区')
 STORED AS ORC
"""
# re.DOTALL 或 re.S - 使 . 匹配包括换行符在内的任意字符。
result = re.match(r"^CREATE.*?\(\n(.*?)\)\n.*",sql1,re.S)
arrs = result.group(1)
print(arrs)
# for arr in arrs:
#     print(arr)