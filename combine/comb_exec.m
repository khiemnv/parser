function ret = comb_execc(cmd,sys,varargin)
    try
    switch(cmd)
        case 'getBaseIn'
            ret = getBaseIn(sys);
        case 'getBaseOut'
            ret = getBaseOut(sys);
        case 'addIn'
            zin = varargin{1};
            pos = varargin{2};
            ret = addIn(sys,zin,pos);
        case 'addOut'
            zout = varargin{1};
            pos = varargin{2};
            ret = addOut(sys,zout,pos);
        case 'addLine'
            src = varargin{1};
            des = varargin{2};
            ret = addLine(sys,src,des);
    end
    catch e
        display(e);
    end
end
function l = addLine(sys,src,des)
    l = add_line(sys,src,des,'autorouting','on');
end
function ret = addIn(sys,zin,pos)
    ret = add_block('simulink/Ports & Subsystems/In1',[sys '/' zin]);
    set_param(ret, 'position', sprintf('[%d %d %d %d]',pos));
end
function ret = addOut(sys,zout,pos)
    ret = add_block('simulink/Ports & Subsystems/Out1',[sys '/' zout]);
    set_param(ret, 'position', sprintf('[%d %d %d %d]',pos));
end
function ret = getBaseIn(sys)
    b = find_system(sys,'searchdepth','1','type','Block');
    pos = get_param(b,'position');
    m = vertcat(pos{:});
    if length(b) == 1
        pos2 = [m(1) m(2)];
    else
        pos2 = [min(m(:,1) - 150, min(m(:,2))];
    end
    d = [40 20];
    ret = [(pos2) (pos2 + d)];
end
function ret = getBaseOut(sys)
    b = find_system(sys,'searchdepth','1','type','Block');
    pos = get_param(b,'position');
    m = vertcat(pos{:});
    if length(b) == 1
        pos2 = [m(3) m(4)];
    else
        pos2 = [min(m(:,3)) - 150, min(m(:,2))];
    end
    d = [40 20];
    ret = [(pos2 - d) (pos2)];
end