local key = KEYS[1]
local lockId = ARGV[1]
-- 当前线程未获取锁
if (redis.call('hexists', key, lockId) == 0) then
  return nil
end
local count = redis.call('hincrby', key, lockId, -1)
if (count > 0) then
  redis.call('pexpire', key, ARGV[2])
  return 0
else
  redis.call('del', key)
  return 1
end
