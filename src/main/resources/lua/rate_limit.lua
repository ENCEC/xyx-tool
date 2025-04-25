local num = redis.call('incr', KEYS[1])
if tonumber(num) == 1 then
    redis.call('expire', KEYS[1], ARGV[1])
    return 1
elseif tonumber(num) > tonumber(ARGV[2]) then
    return 0
else
    return 1
end