import listviewLogo from '../assets/listview.png';

interface LogoProps {
  size?: number;
  className?: string;
}

function Logo({ size = 32, className }: LogoProps) {
  return (
    <img
      src={listviewLogo}
      alt="ListView"
      width={size}
      height={size}
      className={className}
      style={{ objectFit: 'contain' }}
    />
  );
}

export default Logo;
