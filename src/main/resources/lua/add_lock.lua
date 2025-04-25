local key = KEYS[1]
local lockId = ARGV[1]
local expireTime = ARGV[2]
-- 判断锁是否存在
if (redis.call('exists', key) == 0) then
  redis.call('hset', key, lockId, 1)
  redis.call('pexpire', key, expireTime)
  return 1
end
-- 判断是否是当前线程持有锁
if (redis.call('hexists', key, lockId) == 1) then
  -- 如果当前线程已经获取锁了，则进行累加计数器
  redis.call('hincrby', key, lockId, 1)
  -- 重置过期时间
  redis.call('pexpire', key, expireTime)
  return 1
end
return 0
