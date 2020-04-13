import * as React from "react";



const about = () => {
    
    const version = Office.context.mailbox.diagnostics.hostVersion;
  return (
    <div>
      <h3 className="ms-fontSize-16">Office Version: {version}</h3>
      <h3 className="ms-fontSize-16">About this app:</h3>
      <p className="ms-font-m ms-fontColor-neutralPrimary">
        Lorem ipsum dolor sit amet, consectetur adipiscing elit. Duis porttitor diam felis, vel pulvinar leo consequat
        in. Fusce dapibus sed metus tincidunt bibendum. Aliquam erat volutpat. Maecenas vulputate nisl et purus iaculis,
        eu lobortis nisi tempor.
      </p>
    </div>
  );
};

export default about;
