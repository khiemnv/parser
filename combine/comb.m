function comb()
sys = gcs;
baseIn = comb_exec('getBaseIn',sys);
baseOut = comb_exec('getBaseOut',sys);
comb_exec('addIn',sys,'signalA_prev',baseIn + [0 0 0 0]);
comb_exec('addIn',sys,'signalB',baseIn + [0 50 0 50]);
comb_exec('addOut',sys,'signalA',baseOut + [0 0 0 0]);
comb_exec('addOut',sys,'signalC',baseOut + [0 50 0 50]);
comb_exe('addLine',sys,'signalA_prev/1','Cal_signalA/1');
comb_exe('addLine',sys,'Cal_signalA/1','Cal_signalC/1');
comb_exe('addLine',sys,'signalB/1','Cal_signalC/2');
comb_exe('addLine',sys,'Cal_signalA/1','signalA/1');
comb_exe('addLine',sys,'Cal_signalC/1','signalC/1');
