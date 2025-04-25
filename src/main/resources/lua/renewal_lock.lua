local key = KEYS[1]
local lockId = ARGV[1]
local expireTime = ARGV[2]
-- 判断是否是当前线程持有锁
if (redis.call('hexists', key, lockId) == 1) then
  -- 重置过期时间
  redis.call('pexpire', key, expireTime)
  return 1
end
return 0
