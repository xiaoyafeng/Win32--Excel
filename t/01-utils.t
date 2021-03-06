use strict;
use Data::Dumper;
use Test::More;
use Test::Deep;
use Test::Harness;
use Win32::ExcelSimple qw/transpose_array cell2aa aa2cell/;


is(cell2aa(2),    'B',   "convert 2 to B");
is(aa2cell('AB'),   28,  "convert AB to 28");

TODO: {   
	todo_skip "more complex cell function is not finished yet", 3 if 1;


is(aa2cell('AB3'),  [28, 3], "convert AB3 to [28, 3]");
is(cell2aa([2,34]), 'B34', "convert (2,34) to B34");
is(aa2cell('AB3:AB4'), [[28,3], [28,4]], "convert AB3:AB4 to  [[28,3], [28,4]]");
}
cmp_deeply(transpose_array([[1, 'a'],[2,'b']]),  [[1,2],['a','b']], "transpose array");
done_testing;

