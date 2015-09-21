% David Smith's extraction method using matlab and CST V5
% Methods described in: D. R. Smith, D. C. Vier, Th. Koschny, C. M. Soukoulis
% Physical Review E, 71, 036617 (2005)
% and
% D. R. Smith, S. Schultz, P. Markos and C. M. Soukoulis
% Physical Review B, 65, 195104 (2002)
% V3 incorporates Ruopeng Liu's Spatial Dispersion Transformation
% V5 fixes branch errors in real(n)

eps_0 =8.8541878176e-12; % define variables
mu_0 = 1.2566e-6;

% Call CST data

d = sendA(1); % Get slab length from CST
omega= 2*pi*f; % Calculate angular frequency omega
k = omega/3e8; % Calculate wavenumber k
nf = length(s11mag); % Get number of frequency points

s11 = s11mag.*exp(-i.*s11phase*pi/180);
s21 = s21mag.*exp(-i.*s21phase*pi/180);

% invert s-parameters

n_eff = acos((1-(s11.^2-s21.^2)).*(1./(2*s21)))./k/d;

z_eff = sqrt(((1+s11).^2-s21.^2)./((1-s11).^2-s21.^2));

% due to passivity, impose conditions on imag (n) and real (z)

for ii = 1:nf
    if imag(n_eff(ii)) < 0 
       n_eff(ii) = -n_eff(ii);
    end
    if real(z_eff(ii)) < 0 
       z_eff(ii) = -z_eff(ii);
    end
end


eps_eff = n_eff./z_eff;
mu_eff = n_eff.*z_eff;

re_eps_eff = real(eps_eff); %define variables for use in CST Macro
im_eps_eff = imag(eps_eff);
re_mu_eff = real(mu_eff);
im_mu_eff = imag(mu_eff);
re_z_eff = real(z_eff);
im_z_eff = imag(z_eff);
re_n_eff= real (n_eff);
im_n_eff= imag (n_eff);

% Spatial Dispersion

% phase advance

theta = omega*d.*sqrt(mu_eff.*eps_eff*eps_0*mu_0);

magic = tan(theta/2)./(theta/2);

eps_av = eps_eff.*magic;
mu_av = mu_eff.*magic;

re_eps_av = real(eps_av); %define variables for use in CST Macro
im_eps_av = imag(eps_av);
re_mu_av = real(mu_av);
im_mu_av = imag(mu_av);

