#!/bin/sh

set -eu

if [ "$#" -lt 2 ]
then
        echo Usage: `basename $0` "<limit> <command>..."
        exit 1
fi

limit="$1"
shift

cgname="limitmem_$$"
#echo "limiting memory to $limit (cgroup $cgname) for command $@" >&2

cgm create memory "$cgname" >/dev/null
cgm setvalue memory "$cgname" memory.limit_in_bytes "$limit" >/dev/null
# try also limiting swap usage, but this fails if the system has no swap
cgm setvalue memory "$cgname" memory.memsw.limit_in_bytes "$limit" >/dev/null 2>&1 || true
bytes_limit=`cgm getvalue memory "$cgname" memory.limit_in_bytes | tail -1 | cut -f2 -d\"`

# spawn subshell to run in the cgroup
# set +e so a failing child does not prevent us from removing the cgroup
set +e
(
set -e
cgm movepid memory "$cgname" `sh -c 'echo $PPID'` > /dev/null
exec "$@"
)

# grab exit code 
exitcode=`echo $?`

set -e

peak_mem=`cgm getvalue memory "$cgname" memory.max_usage_in_bytes | tail -1 | cut -f2 -d\"`
failcount=`cgm getvalue memory "$cgname" memory.failcnt | tail -1 | cut -f2 -d\"`
percent=`expr "$peak_mem" / \( "$bytes_limit" / 100 \)`
#echo "peak memory used: $peak_mem ($percent%); exceeded limit $failcount times" >&2

cgm remove memory "$cgname" >/dev/null

exit $exitcode

